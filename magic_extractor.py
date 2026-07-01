from __future__ import annotations

import json
import gzip
import os
import pickle
import re
import shutil
import threading
import time
import unicodedata
from dataclasses import dataclass
from pathlib import Path
from tkinter import BooleanVar, END, StringVar, Tk, filedialog, messagebox, ttk

import requests
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter
from PIL import Image, ImageEnhance, ImageGrab, ImageOps, ImageTk

try:
    import pytesseract
except Exception:  # pragma: no cover - handled at runtime for friendly UI
    pytesseract = None

try:
    from rapidfuzz import fuzz, process
except Exception:  # pragma: no cover - difflib fallback keeps the app usable
    fuzz = None
    process = None


APP_DIR = Path(__file__).resolve().parent
APP_VERSION = "2026-07-01.18"
DATA_DIR = APP_DIR / "data"
SCRYFALL_CARDS_PATH = DATA_DIR / "scryfall_default_cards.json"
SETS_PATH = DATA_DIR / "scryfall_sets.json"
MTGJSON_ATOMIC_PATH = DATA_DIR / "mtgjson_atomic_cards.json.gz"
INDEX_PATH = DATA_DIR / "card_index.pkl"
INDEX_VERSION = 3
EXCEL_PATH = APP_DIR / "magic_cards.xlsx"
DEBUG_PATH = APP_DIR / "last_extraction_debug.txt"
HTTP_HEADERS = {
    "User-Agent": "magic-extractor/1.0 (local desktop tool)",
    "Accept": "application/json",
}

FIELDS = [
    "Nome",
    "Tipo",
    "Raridade",
    "Descrição magia",
    "Custo mana",
    "Cor(cores)",
    "Preço minimo",
    "Coleção",
    "É foil?",
    "Ano",
    "Poder",
    "Qtd",
]

RARITY_PT = {
    "common": "Comum",
    "uncommon": "Incomum",
    "rare": "Rara",
    "mythic": "Mítica",
    "special": "Especial",
    "bonus": "Bônus",
}

RARITY_EN = {
    "common": "Common",
    "uncommon": "Uncommon",
    "rare": "Rare",
    "mythic": "Mythic",
    "special": "Special",
    "bonus": "Bonus",
}

COLOR_PT = {
    "W": "Branco",
    "U": "Azul",
    "B": "Preto",
    "R": "Vermelho",
    "G": "Verde",
}

TESSERACT_CANDIDATES = [
    r"C:\Program Files\Tesseract-OCR\tesseract.exe",
    r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
    r"C:\Users\tiago\AppData\Local\Programs\Tesseract-OCR\tesseract.exe",
]


def normalize_text(value: str) -> str:
    value = unicodedata.normalize("NFKD", value or "")
    value = "".join(ch for ch in value if not unicodedata.combining(ch))
    value = value.lower()
    value = re.sub(r"[^a-z0-9]+", " ", value)
    return re.sub(r"\s+", " ", value).strip()


def clean_ocr_line(line: str) -> str:
    line = line.replace("|", "I").strip()
    line = re.sub(r"\s+", " ", line)
    line = re.sub(r"[\{\}\[\]]", "", line)
    return line.strip(" .,:;")


def clean_card_text(value: str) -> str:
    value = value or ""
    value = value.replace("\r\n", "\n").replace("\r", "\n")
    value = re.sub(r"\s*\n\s*", " ", value)
    return re.sub(r"\s+", " ", value).strip()


def collector_key(value: str) -> str:
    value = str(value or "").upper().strip()
    match = re.match(r"0*(\d+)([A-Z]*)$", value)
    if match:
        return f"{int(match.group(1))}{match.group(2)}"
    return value


def ocr_collector_key(value: str) -> str:
    value = str(value or "").upper().strip()
    value = value.translate(str.maketrans({"O": "0", "I": "1", "L": "1"}))
    return collector_key(value)


def fuzzy_word_count(expected_words: list[str], text_words: list[str], threshold: int = 82) -> int:
    count = 0
    for expected in expected_words:
        if any(word == expected or (fuzz and fuzz.ratio(expected, word) >= threshold) for word in text_words):
            count += 1
    return count


def bilingual(en_value: str, pt_value: str) -> str:
    en_value = clean_card_text(en_value)
    pt_value = clean_card_text(pt_value)
    if pt_value and normalize_text(pt_value) != normalize_text(en_value):
        return f"{en_value} / {pt_value}" if en_value else pt_value
    return en_value


def card_faces_text(card: dict, field: str) -> str:
    if card.get(field):
        return card[field]
    faces = card.get("card_faces") or []
    parts = []
    for face in faces:
        value = face.get(field)
        if value:
            name = face.get("name", "")
            parts.append(f"{name}: {value}" if name else value)
    return " // ".join(parts)


def card_power(card: dict) -> str:
    if card.get("power") and card.get("toughness"):
        return f"{card['power']}/{card['toughness']}"
    for face in card.get("card_faces") or []:
        if face.get("power") and face.get("toughness"):
            return f"{face['power']}/{face['toughness']}"
    return ""


def card_mana_cost(card: dict) -> str:
    if card.get("mana_cost"):
        return card["mana_cost"]
    costs = [face.get("mana_cost", "") for face in card.get("card_faces") or []]
    return " // ".join(cost for cost in costs if cost)


def card_colors_pt(card: dict) -> str:
    colors = card.get("colors") or card.get("color_identity") or []
    if not colors:
        return "Incolor"
    return ", ".join(COLOR_PT.get(color, color) for color in colors)


def crop_possible_app_screenshot(image: Image.Image) -> Image.Image:
    width, height = image.size
    if width / max(1, height) <= 1.25:
        return image
    preview_width = min(int(width * 0.45), int(height * 0.62))
    top = int(height * 0.38)
    crop = image.crop((0, top, preview_width, height))
    card_crop = crop_card_preview(crop)
    if card_crop is not None:
        return card_crop
    if crop.width >= 160 and crop.height >= 240:
        return crop
    return image


def crop_card_preview(image: Image.Image) -> Image.Image | None:
    image = image.convert("RGB")
    width, height = image.size
    dark_rows = []
    dark_cols = []
    row_threshold = max(12, int(width * 0.08))
    col_threshold = max(12, int(height * 0.08))

    for y in range(height):
        count = 0
        for x in range(width):
            r, g, b = image.getpixel((x, y))
            if r + g + b < 120:
                count += 1
        if count >= row_threshold:
            dark_rows.append(y)

    for x in range(width):
        count = 0
        for y in range(height):
            r, g, b = image.getpixel((x, y))
            if r + g + b < 120:
                count += 1
        if count >= col_threshold:
            dark_cols.append(x)

    if not dark_rows or not dark_cols:
        return None

    left = max(0, min(dark_cols) - 10)
    right = min(width, max(dark_cols) + 11)
    top = max(0, min(dark_rows) - 10)
    bottom = min(height, max(dark_rows) + 40)
    crop_width = right - left
    crop_height = bottom - top
    if crop_width < 120 or crop_height < 180:
        return None
    ratio = crop_width / max(1, crop_height)
    if not 0.45 <= ratio <= 0.9:
        return None
    return image.crop((left, top, right, bottom))


def configure_tesseract() -> str | None:
    if pytesseract is None:
        return None
    found = shutil.which("tesseract")
    if found:
        pytesseract.pytesseract.tesseract_cmd = found
        return found
    for candidate in TESSERACT_CANDIDATES:
        if Path(candidate).exists():
            pytesseract.pytesseract.tesseract_cmd = candidate
            return candidate
    return None


@dataclass
class OcrResult:
    name_hint: str = ""
    namebar_text: str = ""
    set_code: str = ""
    collector_number: str = ""
    rarity_hint: str = ""
    raw_text: str = ""


class LocalOcr:
    def __init__(self) -> None:
        self.tesseract_path = configure_tesseract()

    def available(self) -> bool:
        return pytesseract is not None and self.tesseract_path is not None

    def extract(self, image: Image.Image) -> OcrResult:
        if not self.available():
            raise RuntimeError(
                "Tesseract OCR não foi encontrado. Instale o Tesseract e reinicie o app."
            )

        image = crop_possible_app_screenshot(image.convert("RGB"))
        texts = []
        namebar_texts = []
        for candidate in self._candidate_images(image):
            width, height = candidate.size
            name_crop = candidate.crop((0, int(height * 0.035), int(width * 0.84), int(height * 0.14)))
            bottom_crop = candidate.crop((0, int(height * 0.86), width, height))
            footer_crop = candidate.crop((0, int(height * 0.74), width, height))
            namebar_texts.extend(
                [
                    self._ocr(name_crop, "--psm 7"),
                    self._ocr(candidate.crop((0, 0, int(width * 0.9), int(height * 0.18))), "--psm 6"),
                ]
            )
            texts.extend(
                [
                    self._ocr(name_crop, "--psm 7"),
                    self._ocr(bottom_crop, "--psm 6"),
                    self._ocr_footer(footer_crop),
                    self._ocr(candidate, "--psm 6"),
                ]
            )

        raw_text = "\n".join(part for part in texts if part.strip())
        namebar_text = "\n".join(part for part in namebar_texts if part.strip())
        name_hint = self._parse_name(raw_text)
        set_code, collector_number, rarity_hint = self._parse_collector(raw_text)
        return OcrResult(
            name_hint=name_hint,
            namebar_text=namebar_text,
            set_code=set_code,
            collector_number=collector_number,
            rarity_hint=rarity_hint,
            raw_text=raw_text,
        )

    @staticmethod
    def _candidate_images(image: Image.Image) -> list[Image.Image]:
        width, height = image.size
        if width / max(1, height) > 0.9:
            candidates = []
            preview_width = min(int(width * 0.45), int(height * 0.62))
            boxes = [
                (0, int(height * 0.38), preview_width, height),
                (0, int(height * 0.45), preview_width, height),
                (0, 0, preview_width, height),
            ]
            for box in boxes:
                crop = image.crop(box)
                if crop.width > 120 and crop.height > 180:
                    candidates.append(crop)
            return candidates
        candidates = [image]
        return candidates

    def _ocr(self, image: Image.Image, config: str) -> str:
        prepared = self._prepare(image)
        return pytesseract.image_to_string(prepared, lang="eng", config=config)

    @staticmethod
    def _ocr_footer(image: Image.Image) -> str:
        footer = image.convert("L")
        footer = footer.resize((footer.width * 5, footer.height * 5))
        simple_binary = footer.point(lambda pixel: 255 if pixel > 100 else 0)
        contrast_footer = ImageEnhance.Contrast(footer).enhance(2.5)
        binary = contrast_footer.point(lambda pixel: 255 if pixel > 100 else 0)
        inverted_binary = ImageOps.invert(contrast_footer).point(lambda pixel: 255 if pixel > 100 else 0)
        return "\n".join(
            [
                pytesseract.image_to_string(simple_binary, lang="eng", config="--psm 6"),
                pytesseract.image_to_string(binary, lang="eng", config="--psm 6"),
                pytesseract.image_to_string(inverted_binary, lang="eng", config="--psm 6"),
            ]
        )

    @staticmethod
    def _prepare(image: Image.Image) -> Image.Image:
        image = image.convert("L")
        scale = max(2, int(1200 / max(1, image.width)))
        image = image.resize((image.width * scale, image.height * scale))
        image = ImageEnhance.Contrast(image).enhance(1.8)
        return image.point(lambda pixel: 255 if pixel > 155 else 0)

    @staticmethod
    def _parse_name(text: str) -> str:
        lines = [clean_ocr_line(line) for line in text.splitlines()]
        lines = [line for line in lines if len(line) >= 3]
        if not lines:
            return ""
        ignored = {
            "colar",
            "imagem",
            "abrir",
            "extrair",
            "foil",
            "copiar",
            "google",
            "docs",
            "cabecalho",
            "salvar",
            "excel",
            "atualizar",
            "base",
            "nome",
            "tipo",
            "raridade",
            "descricao",
            "magia",
            "custo",
            "mana",
            "cor",
            "preco",
            "minimo",
            "colecao",
            "ano",
            "poder",
            "qtd",
            "ocr",
            "erro",
            "changeling",
            "this",
            "card",
            "every",
            "creature",
            "type",
            "is",
        }

        def score(candidate: str) -> int:
            normalized = normalize_text(candidate)
            words = normalized.split()
            if not words or all(word in ignored for word in words):
                return -100
            title_words = re.findall(r"\b[A-Z][a-z]{2,}\b", candidate)
            value = min(len(words), 5)
            value += len(title_words) * 3
            if len(title_words) >= 2:
                value += 10
            elif not title_words:
                value -= 6
            value -= sum(3 for word in words if word in ignored)
            value -= sum(2 for word in words if word in {"instant", "sorcery", "artifact", "enchantment", "land"})
            value -= sum(1 for word in words if word.isdigit())
            return value

        line = max(lines, key=score)
        line = re.sub(r"\s+[\dXWUBRGC/]+$", "", line, flags=re.IGNORECASE)
        return clean_ocr_line(line)

    @staticmethod
    def _parse_collector(text: str) -> tuple[str, str, str]:
        text = text.upper().replace("•", " ")
        text = re.sub(r"[^A-Z0-9]+", " ", text)
        rarity = ""
        rarity_with_pt_match = re.search(
            r"\b([CMUR])\s+(\d{1,4}[A-Z]?)\s+(?:\d+\s+\d+\s+)?([A-Z0-9]{2,5})\s+(?:EN|PT|ES|FR|DE|IT|JP|KO|RU|ZHS|ZHT)\b",
            text,
        )
        if rarity_with_pt_match:
            rarity, number, set_code = rarity_with_pt_match.groups()
            return set_code, number, rarity

        rarity_match = re.search(
            r"\b([CMUR])\s+(\d{1,4}[A-Z]?)\s+([A-Z0-9]{2,5})\s+(?:EN|PT|ES|FR|DE|IT|JP|KO|RU|ZHS|ZHT)\b",
            text,
        )
        if rarity_match:
            rarity, number, set_code = rarity_match.groups()
            return set_code, number, rarity

        match = re.search(r"\b(\d{1,4}[A-Z]?)\s+([A-Z0-9]{2,5})\s+(?:EN|PT|ES|FR|DE|IT|JP|KO|RU|ZHS|ZHT)\b", text)
        if match:
            number, set_code = match.groups()
            return set_code, number, rarity

        return "", "", ""


class ScryfallDatabase:
    def __init__(self) -> None:
        self.cards: list[dict] = []
        self.sets: dict[str, dict] = {}
        self.by_print: dict[tuple[str, str], list[dict]] = {}
        self.pt_by_oracle: dict[str, list[dict]] = {}
        self.en_by_oracle: dict[str, list[dict]] = {}
        self.mtgjson_pt_by_name: dict[str, dict] = {}
        self.name_choices: list[tuple[str, int]] = []
        self.pt_name_choices: list[tuple[str, int]] = []
        self.loaded = False

    def download(self, log) -> None:
        DATA_DIR.mkdir(exist_ok=True)
        log("Buscando metadados de bulk data do Scryfall...")
        bulk = requests.get("https://api.scryfall.com/bulk-data", timeout=60, headers=HTTP_HEADERS)
        bulk.raise_for_status()
        bulk_data = bulk.json()["data"]
        default_cards = next(item for item in bulk_data if item["type"] == "default_cards")

        log("Baixando Default Cards do Scryfall para reconhecer impressões...")
        self._download_file(default_cards["download_uri"], SCRYFALL_CARDS_PATH, log)

        log("Baixando lista de coleções...")
        sets = requests.get("https://api.scryfall.com/sets", timeout=60, headers=HTTP_HEADERS)
        sets.raise_for_status()
        SETS_PATH.write_text(json.dumps(sets.json(), ensure_ascii=False), encoding="utf-8")

        log("Baixando traduções do MTGJSON AtomicCards...")
        self._download_file("https://mtgjson.com/api/v5/AtomicCards.json.gz", MTGJSON_ATOMIC_PATH, log)

        if INDEX_PATH.exists():
            INDEX_PATH.unlink()
        log("Base baixada. Recriando índice local...")
        self.load(force_rebuild=True, log=log)

    @staticmethod
    def _download_file(url: str, path: Path, log) -> None:
        temp_path = path.with_suffix(".tmp")
        with requests.get(url, stream=True, timeout=120, headers=HTTP_HEADERS) as response:
            response.raise_for_status()
            total = int(response.headers.get("content-length", "0") or "0")
            done = 0
            last_log = time.time()
            with temp_path.open("wb") as handle:
                for chunk in response.iter_content(chunk_size=1024 * 1024):
                    if not chunk:
                        continue
                    handle.write(chunk)
                    done += len(chunk)
                    if time.time() - last_log > 2:
                        if total:
                            log(f"Download: {done / total:.0%}")
                        else:
                            log(f"Download: {done // (1024 * 1024)} MB")
                        last_log = time.time()
        temp_path.replace(path)

    def load(self, force_rebuild: bool = False, log=lambda _msg: None) -> None:
        if self.loaded and not force_rebuild:
            return
        if not SCRYFALL_CARDS_PATH.exists():
            raise FileNotFoundError("Base não encontrada. Clique em Atualizar base primeiro.")

        if INDEX_PATH.exists() and not force_rebuild:
            log("Carregando índice local...")
            with INDEX_PATH.open("rb") as handle:
                state = pickle.load(handle)
            if state.get("index_version") != INDEX_VERSION:
                log("Índice antigo detectado. Recriando...")
                return self.load(force_rebuild=True, log=log)
            self.__dict__.update(state)
            self.loaded = True
            return

        log("Lendo arquivo de cartas...")
        with SCRYFALL_CARDS_PATH.open("r", encoding="utf-8") as handle:
            self.cards = json.load(handle)

        if SETS_PATH.exists():
            with SETS_PATH.open("r", encoding="utf-8") as handle:
                sets_payload = json.load(handle)
            self.sets = {item["code"].upper(): item for item in sets_payload.get("data", [])}

        self.by_print.clear()
        self.pt_by_oracle.clear()
        self.en_by_oracle.clear()
        self.mtgjson_pt_by_name.clear()
        self.name_choices.clear()
        self.pt_name_choices.clear()

        seen_choices = set()
        for index, card in enumerate(self.cards):
            set_code = (card.get("set") or "").upper()
            collector = collector_key(card.get("collector_number"))
            if set_code and collector:
                self.by_print.setdefault((set_code, collector), []).append(card)

            oracle_id = card.get("oracle_id") or card.get("name")
            if card.get("lang") == "pt":
                self.pt_by_oracle.setdefault(oracle_id, []).append(card)
            elif card.get("lang") == "en":
                self.en_by_oracle.setdefault(oracle_id, []).append(card)
                for candidate in [card.get("name"), card.get("printed_name")]:
                    key = normalize_text(candidate or "")
                    if key and key not in seen_choices:
                        self.name_choices.append((candidate, index))
                        seen_choices.add(key)

        if MTGJSON_ATOMIC_PATH.exists():
            log("Lendo traduções em português...")
            self.mtgjson_pt_by_name = self._load_mtgjson_portuguese()
            seen_pt_choices = set()
            for index, card in enumerate(self.cards):
                if card.get("lang") != "en":
                    continue
                pt_card = self.mtgjson_pt_by_name.get(normalize_text(card.get("name", "")))
                pt_name = (pt_card or {}).get("printed_name", "")
                key = normalize_text(pt_name)
                if key and key not in seen_pt_choices:
                    self.pt_name_choices.append((pt_name, index))
                    seen_pt_choices.add(key)

        state = {
            "cards": self.cards,
            "sets": self.sets,
            "by_print": self.by_print,
            "pt_by_oracle": self.pt_by_oracle,
            "en_by_oracle": self.en_by_oracle,
            "mtgjson_pt_by_name": self.mtgjson_pt_by_name,
            "name_choices": self.name_choices,
            "pt_name_choices": self.pt_name_choices,
            "loaded": True,
            "index_version": INDEX_VERSION,
        }
        with INDEX_PATH.open("wb") as handle:
            pickle.dump(state, handle, protocol=pickle.HIGHEST_PROTOCOL)
        self.loaded = True

    def find(self, ocr: OcrResult) -> tuple[dict, str]:
        self.load()
        card, set_code, collector_number = self._find_print_in_ocr_text(ocr.raw_text)
        if card:
            return card, f"rodapé OCR {set_code} #{collector_number}"

        namebar_card, namebar_name = self._fuzzy_portuguese_text(ocr.namebar_text)
        if namebar_card:
            namebar_card = self._prefer_print_from_ocr(namebar_card, ocr.raw_text)
            return namebar_card, f"nome superior PT OCR '{namebar_name}'"
        namebar_card = self._fuzzy_name(ocr.namebar_text, strict_words=True)
        if namebar_card:
            namebar_card = self._prefer_print_from_ocr(namebar_card, ocr.raw_text)
            return namebar_card, "nome superior OCR"

        if ocr.set_code and ocr.collector_number:
            candidates = self.by_print.get((ocr.set_code.upper(), collector_key(ocr.collector_number)), [])
            english = [card for card in candidates if card.get("lang") == "en"]
            if english:
                if not ocr.name_hint or self._card_name_matches(english[0], ocr.name_hint):
                    return english[0], f"código {ocr.set_code.upper()} #{ocr.collector_number}"
            if candidates:
                if not ocr.name_hint or self._card_name_matches(candidates[0], ocr.name_hint):
                    return candidates[0], f"código {ocr.set_code.upper()} #{ocr.collector_number}"

        if ocr.name_hint and self._likely_name_hint(ocr.name_hint):
            card = self._fuzzy_name(ocr.name_hint, strict_words=True)
            if card:
                card = self._prefer_print_from_ocr(card, ocr.raw_text)
                return card, f"nome OCR '{ocr.name_hint}'"
        card, matched_name = self._fuzzy_portuguese_text(ocr.raw_text)
        if card:
            card = self._prefer_print_from_ocr(card, ocr.raw_text)
            return card, f"nome PT OCR '{matched_name}'"
        card, matched_line = self._fuzzy_ocr_text(ocr.raw_text)
        if card:
            card = self._prefer_print_from_ocr(card, ocr.raw_text)
            return card, f"texto OCR '{matched_line}'"
        raise LookupError("Não consegui encontrar a carta na base local.")

    @staticmethod
    def _card_name_matches(card: dict, name_hint: str) -> bool:
        card_name = card.get("name", "")
        if not card_name or not name_hint:
            return False
        if normalize_text(card_name) in normalize_text(name_hint):
            return True
        if fuzz:
            return fuzz.WRatio(card_name, name_hint) >= 72
        import difflib

        return difflib.SequenceMatcher(None, normalize_text(card_name), normalize_text(name_hint)).ratio() >= 0.72

    def _prefer_print_from_ocr(self, card: dict, raw_text: str) -> dict:
        oracle_id = card.get("oracle_id")
        raw_upper = raw_text.upper()
        same_oracle_cards = [
            item
            for item in self.cards
            if item.get("lang") == "en" and item.get("oracle_id") == oracle_id
        ]
        for item in same_oracle_cards:
            set_code = (item.get("set") or "").upper()
            if set_code and re.search(rf"\b{re.escape(set_code)}\b", raw_upper):
                return item

        collectors = self._collector_numbers_in_text(raw_text)
        if not collectors:
            return sorted(same_oracle_cards, key=lambda item: item.get("released_at") or "9999-99-99")[0] if same_oracle_cards else card
        current_collector = collector_key(card.get("collector_number"))
        if current_collector in collectors:
            return card
        same_cards = [
            item
            for item in self.cards
            if item.get("lang") == "en"
            and item.get("oracle_id") == oracle_id
            and collector_key(item.get("collector_number")) in collectors
        ]
        if same_cards:
            for item in same_cards:
                if (item.get("set") or "").upper() in raw_upper:
                    return item
            return same_cards[0]
        return sorted(same_oracle_cards, key=lambda item: item.get("released_at") or "9999-99-99")[0] if same_oracle_cards else card

    @staticmethod
    def _collector_numbers_in_text(raw_text: str) -> set[str]:
        text = raw_text.upper()
        collectors = set()
        for match in re.finditer(r"\b([0-9OIL]{1,4}[A-Z]?)\s*/\s*[0-9OILSE]{1,5}\b", text):
            collectors.add(ocr_collector_key(match.group(1)))
        tokens = re.findall(r"[A-Z0-9]+", text)
        for index, token in enumerate(tokens[:-1]):
            normalized_token = token.translate(str.maketrans({"O": "0", "I": "1", "L": "1"}))
            normalized_next = tokens[index + 1].translate(str.maketrans({"O": "0", "I": "1", "L": "1", "E": "8", "S": "5"}))
            if re.fullmatch(r"0*\d{1,4}", normalized_token) and re.fullmatch(r"\d{1,5}", normalized_next):
                total = int(normalized_next)
                number_match = re.match(r"0*(\d+)", normalized_token)
                if number_match and int(number_match.group(1)) <= total:
                    collectors.add(collector_key(normalized_token))
        return collectors

    def _find_print_in_ocr_text(self, raw_text: str) -> tuple[dict | None, str, str]:
        if not raw_text:
            return None, "", ""
        known_sets = {set_code for set_code, _collector in self.by_print.keys()}
        language_codes = {"EN", "PT", "ES", "FR", "DE", "IT", "JP", "KO", "RU", "ZHS", "ZHT"}
        rarity_codes = {"C", "U", "R", "M"}
        tokens = re.findall(r"[A-Z0-9]+", raw_text.upper())

        def corrected_set(token: str) -> str:
            if token in known_sets:
                return token
            ocr_set_fixes = {
                "FRE": "FRF",
                "ECT": "ECL",
                "ECI": "ECL",
                "EC1": "ECL",
            }
            fixed = ocr_set_fixes.get(token, "")
            return fixed if fixed in known_sets else ""

        def candidate_from(set_code: str, collector: str) -> dict | None:
            candidates = [card for card in self.by_print.get((set_code, collector), []) if card.get("lang") == "en"]
            candidates = candidates or self.by_print.get((set_code, collector), [])
            for card in candidates:
                if self._card_name_in_raw_text(card, raw_text):
                    return card
            return None

        for index, token in enumerate(tokens):
            if token not in rarity_codes:
                continue
            collector_index = index + 1
            if collector_index < len(tokens) and tokens[collector_index] in rarity_codes:
                collector_index += 1
            if collector_index >= len(tokens) or not re.fullmatch(r"0*\d{1,4}[A-Z]?", tokens[collector_index]):
                continue
            collector = ocr_collector_key(tokens[collector_index])
            for lookahead in range(collector_index + 1, min(collector_index + 25, len(tokens) - 1)):
                set_code = corrected_set(tokens[lookahead])
                if not set_code:
                    continue
                if tokens[lookahead + 1] not in language_codes and (set_code, collector) not in self.by_print:
                    continue
                card = candidate_from(set_code, collector)
                if card:
                    return card, set_code, collector

        checked = set()
        for index, token in enumerate(tokens):
            if not re.fullmatch(r"0*\d{1,4}[A-Z]?", token):
                continue
            if token.isdigit() and len(token) <= 2:
                previous_is_number = index > 0 and tokens[index - 1].isdigit()
                next_is_number = index + 1 < len(tokens) and tokens[index + 1].isdigit()
                if previous_is_number or next_is_number:
                    continue
            collector = ocr_collector_key(token)
            if index + 1 < len(tokens) and tokens[index + 1].isdigit():
                search_start = index + 2
            else:
                search_start = index + 1
            for lookahead in range(search_start, min(search_start + 25, len(tokens) - 1)):
                set_code = corrected_set(tokens[lookahead])
                if not set_code:
                    continue
                key = (set_code, collector)
                if key in checked:
                    continue
                checked.add(key)
                card = candidate_from(set_code, collector)
                if card:
                    return card, set_code, collector
        return None, "", ""

    def _card_name_in_raw_text(self, card: dict, raw_text: str) -> bool:
        names = [card.get("name", "")]
        pt_card = self.portuguese_for(card)
        if pt_card:
            names.append(pt_card.get("printed_name") or pt_card.get("name") or "")
        names = [name for name in names if name]
        if not names:
            return False
        normalized_raw = normalize_text(raw_text)
        for card_name in names:
            normalized_name = normalize_text(card_name)
            if normalized_name and normalized_name in normalized_raw:
                return True
            name_words = [word for word in normalized_name.split() if len(word) >= 4]
            required_matches = 2 if len(name_words) >= 2 else 1
            raw_words = [word for word in normalized_raw.split() if len(word) >= 3]
            if name_words:
                if fuzzy_word_count(name_words, raw_words) >= required_matches:
                    return True
                joined_raw = " ".join(raw_words)
                joined_name = " ".join(name_words)
                if not fuzz or fuzz.partial_ratio(joined_name, joined_raw) < 78:
                    continue
            if fuzz:
                for line in raw_text.splitlines():
                    cleaned = clean_ocr_line(line)
                    normalized_line = normalize_text(cleaned)
                    line_words = normalized_line.split()
                    if len(line_words) > 8:
                        continue
                    if name_words and fuzzy_word_count(name_words, [word for word in line_words if len(word) >= 3]) < required_matches:
                        continue
                    if fuzz.WRatio(card_name, cleaned) >= 84:
                        return True
        return False

    def is_confident_match(self, card: dict, ocr: OcrResult) -> bool:
        if self._card_name_in_raw_text(card, ocr.raw_text):
            return True
        set_code = (card.get("set") or "").upper()
        collector = collector_key(card.get("collector_number"))
        tokens = re.findall(r"[A-Z0-9]+", ocr.raw_text.upper())
        for index, token in enumerate(tokens):
            if collector_key(token) != collector:
                continue
            window = set(tokens[index + 1 : index + 16])
            if set_code in window:
                return True
        card_type = normalize_text(card_faces_text(card, "type_line"))
        card_text = normalize_text(card_faces_text(card, "oracle_text"))
        raw = normalize_text(ocr.raw_text)
        type_words = [word for word in card_type.split() if len(word) >= 5]
        text_words = [word for word in card_text.split() if len(word) >= 7]
        return (
            sum(1 for word in type_words if word in raw) >= 2
            and sum(1 for word in text_words if word in raw) >= 4
        )

    @staticmethod
    def _likely_name_hint(name_hint: str) -> bool:
        normalized = normalize_text(name_hint)
        words = normalized.split()
        if len(words) < 1 or len(words) > 6:
            return False
        rules_words = {
            "changeling",
            "this",
            "card",
            "every",
            "creature",
            "type",
            "whenever",
            "target",
            "until",
            "magic",
            "gathering",
            "colecao",
            "raridade",
            "descricao",
            "sacrifique",
            "sacrifice",
            "causa",
            "dano",
            "damage",
            "ataca",
            "ataque",
            "bloquear",
        }
        hard_action_words = {
            "sacrifique",
            "sacrifice",
            "causa",
            "dano",
            "damage",
            "ataca",
            "ataque",
            "bloquear",
        }
        if hard_action_words.intersection(words):
            return False
        return len(rules_words.intersection(words)) <= 1

    def _fuzzy_name(self, name_hint: str, strict_words: bool = False) -> dict | None:
        normalized_hint = normalize_text(name_hint)
        if not normalized_hint:
            return None
        if process and fuzz:
            pt_choices = {name: index for name, index in self.pt_name_choices}
            pt_match = process.extractOne(name_hint, pt_choices.keys(), scorer=fuzz.WRatio) if pt_choices else None
            if pt_match and pt_match[1] >= 84:
                return self.cards[pt_choices[pt_match[0]]]
            choices = {name: index for name, index in self.name_choices}
            match = process.extractOne(name_hint, choices.keys(), scorer=fuzz.WRatio)
            if match and match[1] >= 84:
                if strict_words and not self._name_words_present(match[0], normalized_hint):
                    return None
                return self.cards[choices[match[0]]]
            return None

        import difflib

        names = [name for name, _index in self.name_choices]
        match = difflib.get_close_matches(name_hint, names, n=1, cutoff=0.72)
        if match:
            if strict_words and not self._name_words_present(match[0], normalized_hint):
                return None
            index = next(index for name, index in self.name_choices if name == match[0])
            return self.cards[index]
        pt_names = [name for name, _index in self.pt_name_choices]
        pt_match = difflib.get_close_matches(name_hint, pt_names, n=1, cutoff=0.72)
        if pt_match:
            index = next(index for name, index in self.pt_name_choices if name == pt_match[0])
            return self.cards[index]
        return None

    @staticmethod
    def _name_words_present(card_name: str, normalized_text: str) -> bool:
        words = [word for word in normalize_text(card_name).split() if len(word) >= 4]
        if not words:
            return False
        required = 2 if len(words) >= 2 else 1
        return sum(1 for word in words if word in normalized_text) >= required

    def _fuzzy_portuguese_text(self, raw_text: str) -> tuple[dict | None, str]:
        if not raw_text or not process or not fuzz or not self.pt_name_choices:
            return None, ""
        raw_words = [word for word in normalize_text(raw_text).split() if len(word) >= 3]
        best_word_card = None
        best_word_name = ""
        best_word_score = 0
        for pt_name, index in self.pt_name_choices:
            pt_words = [word for word in normalize_text(pt_name).split() if len(word) >= 4]
            if not pt_words:
                continue
            matches = fuzzy_word_count(pt_words, raw_words)
            required = min(2, len(pt_words))
            if matches < required:
                continue
            score = matches * 100 + min(len(pt_words), 6)
            if score > best_word_score:
                best_word_score = score
                best_word_name = pt_name
                best_word_card = self.cards[index]
        if best_word_card:
            return best_word_card, best_word_name

        choices = {name: index for name, index in self.pt_name_choices}
        best_card = None
        best_name = ""
        best_score = 0
        for line in raw_text.splitlines():
            normalized = normalize_text(clean_ocr_line(line))
            words = [word for word in normalized.split() if len(word) >= 3]
            if not words:
                continue
            for size in range(min(5, len(words)), 1, -1):
                for start in range(0, len(words) - size + 1):
                    candidate = " ".join(words[start : start + size])
                    match = process.extractOne(candidate, choices.keys(), scorer=fuzz.WRatio)
                    if not match:
                        continue
                    pt_name = match[0]
                    pt_words = [word for word in normalize_text(pt_name).split() if len(word) >= 4]
                    if len(pt_words) == 1 and len(words) > 2:
                        continue
                    if pt_words and fuzzy_word_count(pt_words, words) < min(2, len(pt_words)):
                        continue
                    adjusted_score = match[1] + (len(pt_words) * 3)
                    if adjusted_score > best_score:
                        best_score = adjusted_score
                        best_name = pt_name
                        best_card = self.cards[choices[pt_name]]
        if best_card and best_score >= 90:
            return best_card, best_name
        return None, ""

    def _fuzzy_ocr_text(self, raw_text: str) -> tuple[dict | None, str]:
        if not raw_text or not process or not fuzz:
            return None, ""
        ignored = {
            "colar",
            "imagem",
            "abrir",
            "extrair",
            "foil",
            "copiar",
            "google",
            "docs",
            "cabecalho",
            "salvar",
            "excel",
            "atualizar",
            "base",
            "nome",
            "tipo",
            "raridade",
            "descricao",
            "magia",
            "custo",
            "mana",
            "cor",
            "preco",
            "minimo",
            "colecao",
            "ano",
            "poder",
            "qtd",
            "ocr",
            "erro",
            "changeling",
            "this",
            "card",
            "every",
            "creature",
            "type",
            "is",
        }
        choices = {name: index for name, index in self.name_choices}
        best_card = None
        best_line = ""
        best_score = 0
        for line in raw_text.splitlines():
            cleaned = clean_ocr_line(line)
            normalized = normalize_text(cleaned)
            words = [word for word in normalized.split() if word not in ignored and len(word) > 1]
            if len(words) < 2:
                continue
            for size in range(min(5, len(words)), 1, -1):
                for start in range(0, len(words) - size + 1):
                    candidate = " ".join(words[start : start + size])
                    if len(candidate) < 6 or not any(len(word) >= 4 for word in candidate.split()):
                        continue
                    match = process.extractOne(candidate, choices.keys(), scorer=fuzz.WRatio)
                    if match and match[1] > best_score:
                        best_score = match[1]
                        best_card = self.cards[choices[match[0]]]
                        best_line = cleaned
        if best_card and best_score >= 88:
            return best_card, best_line
        return None, ""

    @staticmethod
    def _load_mtgjson_portuguese() -> dict[str, dict]:
        with gzip.open(MTGJSON_ATOMIC_PATH, "rt", encoding="utf-8") as handle:
            payload = json.load(handle)
        translations = {}
        for english_name, variants in payload.get("data", {}).items():
            if isinstance(variants, dict):
                variants = [variants]
            for variant in variants or []:
                foreign_items = variant.get("foreignData") or []
                portuguese = next(
                    (
                        item
                        for item in foreign_items
                        if item.get("language") in {"Portuguese", "Português", "Portuguese (Brazil)"}
                    ),
                    None,
                )
                if portuguese:
                    translations[normalize_text(english_name)] = {
                        "printed_name": portuguese.get("name", ""),
                        "printed_type_line": portuguese.get("type", ""),
                        "printed_text": portuguese.get("text", ""),
                    }
                    break
        return translations

    def portuguese_for(self, card: dict) -> dict | None:
        oracle_id = card.get("oracle_id") or card.get("name")
        candidates = self.pt_by_oracle.get(oracle_id, [])
        if not candidates:
            return self.mtgjson_pt_by_name.get(normalize_text(card.get("name", "")))
        same_set = [item for item in candidates if item.get("set") == card.get("set")]
        return same_set[0] if same_set else candidates[0]

    def to_row(self, card: dict, foil: bool) -> list[str]:
        pt_card = self.portuguese_for(card)
        rarity = card.get("rarity", "")
        rarity_value = bilingual(RARITY_EN.get(rarity, rarity.title()), RARITY_PT.get(rarity, ""))

        pt_name = (pt_card or {}).get("printed_name") or (pt_card or {}).get("name") or ""
        pt_type = (pt_card or {}).get("printed_type_line") or (pt_card or {}).get("type_line") or ""
        pt_text = (pt_card or {}).get("printed_text") or card_faces_text(pt_card or {}, "printed_text")

        set_code = (card.get("set") or "").upper()
        set_info = self.sets.get(set_code, {})
        set_name = card.get("set_name") or set_info.get("name") or ""
        released_at = card.get("released_at") or set_info.get("released_at") or ""
        finishes = card.get("finishes") or []
        foil_value = "Sim" if foil or finishes == ["foil"] else "Não"

        row = [
            bilingual(card.get("name", ""), pt_name),
            bilingual(card_faces_text(card, "type_line"), pt_type),
            rarity_value,
            bilingual(card_faces_text(card, "oracle_text"), pt_text),
            card_mana_cost(card),
            card_colors_pt(card),
            "",
            f"{set_name} ({set_code})" if set_name else set_code,
            foil_value,
            released_at[:4] if released_at else "",
            card_power(card),
            "1",
        ]
        return row


class MagicExtractorApp:
    def __init__(self, root: Tk) -> None:
        self.root = root
        self.main_thread_id = threading.get_ident()
        self.root.title(f"Magic Extractor v{APP_VERSION}")
        self.root.geometry("1180x760")
        self.db = ScryfallDatabase()
        self.ocr = LocalOcr()
        self.current_image: Image.Image | None = None
        self.preview_ref = None
        self.busy_widgets = []
        self.extraction_id = 0
        self.field_vars = {field: StringVar() for field in FIELDS}
        self.foil_var = BooleanVar(value=False)
        self.status_var = StringVar(value=f"Pronto. v{APP_VERSION}")
        self.raw_ocr_var = StringVar(value="")
        self._build_ui()
        self._log_ocr_status()

    def _build_ui(self) -> None:
        self.root.columnconfigure(1, weight=1)
        self.root.rowconfigure(0, weight=1)

        left = ttk.Frame(self.root, padding=10)
        left.grid(row=0, column=0, sticky="ns")
        left.rowconfigure(2, weight=1)

        buttons = ttk.Frame(left)
        buttons.grid(row=0, column=0, sticky="ew")

        def add_button(row: int, text: str, command, pady=2):
            button = ttk.Button(buttons, text=text, command=command)
            button.grid(row=row, column=0, sticky="ew", pady=pady)
            self.busy_widgets.append(button)
            return button

        add_button(0, "Colar imagem", self.paste_image)
        add_button(1, "Abrir imagem", self.open_image)
        self.extract_button = add_button(2, "Extrair", self.extract)
        foil_check = ttk.Checkbutton(buttons, text="Foil", variable=self.foil_var)
        foil_check.grid(row=3, column=0, sticky="w", pady=4)
        self.busy_widgets.append(foil_check)
        add_button(4, "Copiar linha", self.copy_row)
        add_button(5, "Copiar Google Docs", self.copy_google_docs)
        add_button(6, "Copiar com cabeçalho", self.copy_with_header)
        add_button(7, "Salvar Excel", self.save_excel)
        add_button(8, "Atualizar base", self.download_database, pady=12)

        self.progress = ttk.Progressbar(left, mode="indeterminate")
        self.progress.grid(row=1, column=0, sticky="ew", pady=(0, 8))
        self.progress.grid_remove()

        self.preview = ttk.Label(left, text="Cole ou abra uma imagem", anchor="center")
        self.preview.grid(row=2, column=0, sticky="nsew", pady=10)

        right = ttk.Frame(self.root, padding=10)
        right.grid(row=0, column=1, sticky="nsew")
        right.columnconfigure(1, weight=1)

        for row_index, field in enumerate(FIELDS):
            ttk.Label(right, text=field).grid(row=row_index, column=0, sticky="nw", padx=(0, 8), pady=3)
            entry = ttk.Entry(right, textvariable=self.field_vars[field])
            if field == "Descrição magia":
                entry = ttk.Entry(right, textvariable=self.field_vars[field])
            entry.grid(row=row_index, column=1, sticky="ew", pady=3)

        ttk.Label(right, text="OCR bruto").grid(row=len(FIELDS), column=0, sticky="nw", padx=(0, 8), pady=3)
        self.raw_text = ttk.Label(right, textvariable=self.raw_ocr_var, wraplength=850, justify="left")
        self.raw_text.grid(row=len(FIELDS), column=1, sticky="ew", pady=3)

        status = ttk.Label(self.root, textvariable=self.status_var, anchor="w", padding=(10, 4))
        status.grid(row=1, column=0, columnspan=2, sticky="ew")

    def _log_ocr_status(self) -> None:
        if self.ocr.available():
            self.log(f"OCR pronto: {self.ocr.tesseract_path} | v{APP_VERSION}")
        else:
            self.log(f"OCR indisponível: instale o Tesseract para extrair texto da imagem. | v{APP_VERSION}")

    def log(self, message: str) -> None:
        if threading.get_ident() != self.main_thread_id:
            self.root.after(0, lambda: self.log(message))
            return
        self.status_var.set(message)
        self.root.update_idletasks()

    def set_busy(self, busy: bool, message: str = "") -> None:
        if busy:
            if message:
                self.status_var.set(message)
            self.progress.grid()
            self.progress.start(12)
            self.root.configure(cursor="watch")
            state = "disabled"
        else:
            self.progress.stop()
            self.progress.grid_remove()
            self.root.configure(cursor="")
            state = "normal"
        for widget in self.busy_widgets:
            widget.configure(state=state)
        self.root.update_idletasks()

    def clear_results(self) -> None:
        for field in FIELDS:
            self.field_vars[field].set("")
        self.raw_ocr_var.set("")

    def run_background(self, target, done_message: str | None = None, job_id: int | None = None) -> None:
        def is_current_job() -> bool:
            return job_id is None or job_id == self.extraction_id

        def worker() -> None:
            try:
                target()
                if done_message and is_current_job():
                    self.root.after(0, lambda: self.log(done_message))
            except Exception as exc:
                message = str(exc)
                if is_current_job():
                    self.root.after(0, self.clear_results)
                    self.root.after(0, lambda: self.raw_ocr_var.set(f"Erro: {message}"))
                    self.root.after(0, lambda: messagebox.showerror("Erro", message))
                    self.root.after(0, lambda: self.log(f"Erro: {message}"))
            finally:
                if is_current_job():
                    self.root.after(0, lambda: self.set_busy(False))

        threading.Thread(target=worker, daemon=True).start()

    def paste_image(self) -> None:
        grabbed = ImageGrab.grabclipboard()
        if isinstance(grabbed, Image.Image):
            self.set_image(grabbed)
            return
        if isinstance(grabbed, list) and grabbed:
            self.set_image(Image.open(grabbed[0]))
            return
        messagebox.showwarning("Sem imagem", "Não encontrei imagem na área de transferência.")

    def open_image(self) -> None:
        file_path = filedialog.askopenfilename(
            title="Abrir imagem",
            filetypes=[("Imagens", "*.png;*.jpg;*.jpeg;*.webp;*.bmp"), ("Todos", "*.*")],
        )
        if file_path:
            self.set_image(Image.open(file_path))

    def set_image(self, image: Image.Image) -> None:
        self.current_image = crop_possible_app_screenshot(image.convert("RGB"))
        self.clear_results()
        preview = self.current_image.copy()
        preview.thumbnail((340, 520))
        self.preview_ref = ImageTk.PhotoImage(preview)
        self.preview.configure(image=self.preview_ref, text="")
        self.log("Imagem carregada.")

    def refresh_image_from_clipboard(self) -> bool:
        grabbed = ImageGrab.grabclipboard()
        image = None
        if isinstance(grabbed, Image.Image):
            image = grabbed
        elif isinstance(grabbed, list) and grabbed:
            try:
                image = Image.open(grabbed[0])
            except Exception:
                image = None
        if image is None:
            return False
        self.current_image = crop_possible_app_screenshot(image.convert("RGB"))
        preview = self.current_image.copy()
        preview.thumbnail((340, 520))
        self.preview_ref = ImageTk.PhotoImage(preview)
        self.preview.configure(image=self.preview_ref, text="")
        return True

    def extract(self) -> None:
        if self.current_image is None:
            messagebox.showwarning("Sem imagem", "Cole ou abra uma imagem primeiro.")
            return
        self.extraction_id += 1
        job_id = self.extraction_id
        image_for_job = self.current_image.copy()
        self.clear_results()
        self.raw_ocr_var.set(f"Extraindo dados... v{APP_VERSION}")
        self.set_busy(True, "Extraindo dados da carta...")
        try:
            self.log("Carregando base local...")
            self.db.load(log=self.log)
            self.log("Rodando OCR local...")
            ocr_result = self.ocr.extract(image_for_job)
            self.log("Procurando carta na base...")
            card, reason = self.db.find(ocr_result)
            self.write_debug(card, reason, ocr_result)
            if not self.db.is_confident_match(card, ocr_result):
                raise RuntimeError(
                    f"Resultado suspeito ({card.get('name')}). Campos foram limpos; veja {DEBUG_PATH.name}."
                )
            row = self.db.to_row(card, self.foil_var.get())
            self.fill_fields(row, ocr_result, reason, job_id)
        except Exception as exc:
            message = str(exc)
            self.clear_results()
            self.raw_ocr_var.set(f"Erro: {message}")
            self.log(f"Erro: {message}")
            messagebox.showerror("Erro", message)
        finally:
            self.set_busy(False)

    def write_debug(self, card: dict, reason: str, ocr_result: OcrResult) -> None:
        lines = [
            f"version: {APP_VERSION}",
            f"matched_name: {card.get('name', '')}",
            f"matched_set: {(card.get('set') or '').upper()}",
            f"matched_collector: {card.get('collector_number', '')}",
            f"reason: {reason}",
            f"name_hint: {ocr_result.name_hint}",
            f"namebar_text: {ocr_result.namebar_text}",
            f"set_code_hint: {ocr_result.set_code}",
            f"collector_hint: {ocr_result.collector_number}",
            "",
            "raw_ocr:",
            ocr_result.raw_text,
        ]
        DEBUG_PATH.write_text("\n".join(lines), encoding="utf-8")

    def fill_fields(self, row: list[str], ocr_result: OcrResult, reason: str, job_id: int) -> None:
        if job_id != self.extraction_id:
            return
        self.clear_results()
        for field, value in zip(FIELDS, row):
            self.field_vars[field].set(value)
        raw = ocr_result.raw_text.strip()
        self.raw_ocr_var.set(raw[:1200])
        self.log(f"Preenchido: {row[0]} | {reason}")

    def row_values(self) -> list[str]:
        return [self.field_vars[field].get() for field in FIELDS]

    def copy_row(self) -> None:
        self.root.clipboard_clear()
        self.root.clipboard_append("\t".join(self.row_values()))
        self.log("Linha copiada em formato Google Sheets.")

    def copy_google_docs(self) -> None:
        value = " | ".join(self.row_values())
        self.root.clipboard_clear()
        self.root.clipboard_append(value)
        self.log("Linha copiada em texto simples para Google Docs.")

    def copy_with_header(self) -> None:
        value = "\t".join(FIELDS) + "\n" + "\t".join(self.row_values())
        self.root.clipboard_clear()
        self.root.clipboard_append(value)
        self.log("Cabeçalho e linha copiados.")

    def save_excel(self) -> None:
        row = self.row_values()
        if not any(row):
            messagebox.showwarning("Sem dados", "Extraia ou preencha uma carta primeiro.")
            return
        if EXCEL_PATH.exists():
            workbook = load_workbook(EXCEL_PATH)
            sheet = workbook.active
        else:
            workbook = Workbook()
            sheet = workbook.active
            sheet.title = "Cartas"
            sheet.append(FIELDS)
            for cell in sheet[1]:
                cell.font = Font(bold=True)
                cell.fill = PatternFill("solid", fgColor="D9EAF7")
            for index, field in enumerate(FIELDS, start=1):
                sheet.column_dimensions[get_column_letter(index)].width = max(12, min(42, len(field) + 8))
        sheet.append(row)
        workbook.save(EXCEL_PATH)
        self.log(f"Salvo em {EXCEL_PATH}")

    def download_database(self) -> None:
        self.run_background(lambda: self.db.download(self.log), "Base atualizada.")


def main() -> None:
    root = Tk()
    app = MagicExtractorApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
