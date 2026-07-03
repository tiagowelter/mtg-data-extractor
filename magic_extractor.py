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
from tkinter import BooleanVar, END, Label, StringVar, Tk, filedialog, messagebox, ttk

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


def configure_ssl_certificates() -> None:
    import certifi

    default_bundle = certifi.where()
    for var in ("SSL_CERT_FILE", "REQUESTS_CA_BUNDLE", "CURL_CA_BUNDLE"):
        value = os.environ.get(var)
        if value and not Path(value).is_file():
            os.environ[var] = default_bundle


configure_ssl_certificates()


APP_DIR = Path(__file__).resolve().parent
APP_VERSION = "2026-07-03.2"
DATA_DIR = APP_DIR / "data"
SCRYFALL_CARDS_PATH = DATA_DIR / "scryfall_default_cards.json"
SETS_PATH = DATA_DIR / "scryfall_sets.json"
MTGJSON_ATOMIC_PATH = DATA_DIR / "mtgjson_atomic_cards.json.gz"
INDEX_PATH = DATA_DIR / "card_index.pkl"
INDEX_VERSION = 9
# Layouts that reuse real card names but are never scanned as collectible cards;
# excluded from the fuzzy name/face matching pools.
_NON_CATALOG_LAYOUTS = {"art_series", "token", "double_faced_token", "emblem"}
EXCEL_PATH = APP_DIR / "MagicCollection.xlsx"
DEBUG_PATH = APP_DIR / "last_extraction_debug.txt"
DEBUG_LOG_PATH = APP_DIR / "extraction_log.txt"
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
    str(Path(os.environ.get("LOCALAPPDATA", "")) / "Programs" / "Tesseract-OCR" / "tesseract.exe"),
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


COMMON_PT_OCR_WORDS = {
    "cemiterio",
    "cemiterios",
    "grimorio",
    "grimorios",
    "conjura",
    "conjurar",
    "magica",
    "magicas",
    "criatura",
    "criaturas",
    "terreno",
    "terrenos",
    "oponente",
    "oponentes",
    "jogador",
    "jogadores",
    "card",
    "cards",
    "turno",
    "turnos",
    "poder",
    "resistencia",
    "controle",
    "alvo",
    "feiticaria",
    "encantamento",
    "instantanea",
    "instantaneas",
    "revela",
    "revelar",
    "coloque",
    "comprar",
    "compre",
    "exila",
    "exile",
    "dano",
    "combate",
    "etapa",
    "final",
    "inicio",
    "proprio",
    "propria",
    "numero",
    "cada",
    "vezes",
    "vez",
    "quando",
    "voce",
    "seus",
    "suas",
    "seu",
    "sua",
    "ate",
    "todo",
    "toda",
    "todos",
    "todas",
    "com",
    "para",
    "mais",
    "menos",
    "outro",
    "outra",
    "outros",
    "outras",
}

COMMON_EN_OCR_WORDS = {
    "card",
    "cards",
    "draw",
    "land",
    "lands",
    "damage",
    "creature",
    "creatures",
    "target",
    "opponent",
    "opponents",
    "control",
    "enters",
    "whenever",
    "ability",
    "turn",
    "spell",
    "spells",
    "player",
    "players",
    "graveyard",
    "library",
    "battlefield",
    "permanent",
    "permanents",
    "counter",
    "counters",
    "token",
    "tokens",
    "serving",
    "nothing",
    "like",
    "posters",
    "made",
    "second",
    "resolved",
    "legendary",
    "pilot",
}

AMBIGUOUS_SET_CODE_WORDS = {
    "A",
    "AND",
    "ARE",
    "AS",
    "COM",
    "DAS",
    "DE",
    "DOS",
    "E",
    "EM",
    "MED",
    "NEM",
    "NO",
    "POR",
    "STA",
    "THE",
    "UMA",
}


def is_plausible_namebar(text: str) -> bool:
    cleaned = clean_ocr_line(text or "")
    words = re.findall(r"[A-Za-zÀ-ü]{3,}", cleaned)
    if not words or len(words) > 6:
        return False
    if not any(len(word) >= 4 for word in words):
        return False
    alpha = sum(ch.isalpha() for ch in cleaned)
    if cleaned and alpha < len(cleaned) * 0.35:
        return False
    title_words = re.findall(r"\b[A-ZÀ-Ü][a-zà-ü]{2,}\b", cleaned)
    if not title_words and len(words) <= 3:
        return False
    return True


def _namebar_line_score(line: str) -> int:
    score = sum(1 for word in line.split() if len(word) >= 4)
    if "," in line:
        score += 5
    if re.match(r"^[A-ZÀ-Ü]", line):
        score += 2
    return score


def ranked_namebar_lines(text: str) -> list[str]:
    lines = [clean_ocr_line(line) for line in (text or "").splitlines()]
    candidates = [line for line in lines if is_plausible_namebar(line)]
    ranked = sorted(candidates, key=_namebar_line_score, reverse=True)
    seen = set()
    unique = []
    for line in ranked:
        key = normalize_text(line)
        if key and key not in seen:
            seen.add(key)
            unique.append(line)
    return unique


def best_namebar_line(text: str) -> str:
    ranked = ranked_namebar_lines(text)
    return ranked[0] if ranked else ""


def normalize_ocr_name(value: str) -> str:
    value = clean_ocr_line(value or "")
    tokens = []
    for token in value.split():
        if re.search(r"\d", token) and re.search(r"[A-Za-z]", token):
            token = token.translate(str.maketrans({"1": "l", "0": "o", "5": "s", "8": "b"}))
        tokens.append(token)
    return " ".join(tokens)


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


def words_fuzzy_match(expected: str, word: str, threshold: int = 82) -> bool:
    if expected == word:
        return True
    if not fuzz:
        return False
    length_gap = abs(len(expected) - len(word))
    effective_threshold = threshold
    if length_gap >= 3:
        effective_threshold = max(threshold, 90)
    elif length_gap >= 2:
        effective_threshold = max(threshold, 88)
    ratio = fuzz.ratio(expected, word)
    if ratio < effective_threshold:
        return False
    if length_gap >= 3 and min(len(expected), len(word)) >= 5:
        shorter = expected if len(expected) < len(word) else word
        longer = word if shorter == expected else expected
        if longer.startswith(shorter[: max(4, len(shorter))]):
            return ratio >= 92
    return True


def fuzzy_word_count(expected_words: list[str], text_words: list[str], threshold: int = 82) -> int:
    count = 0
    for expected in expected_words:
        if any(words_fuzzy_match(expected, word, threshold) for word in text_words):
            count += 1
    return count


def detect_tesseract_languages() -> tuple[str, bool]:
    if pytesseract is None:
        return "eng", False
    try:
        langs = pytesseract.get_languages(config="")
        if "por" in langs:
            return "por+eng", True
    except Exception:
        pass
    return "eng", False


def bilingual(en_value: str, pt_value: str) -> str:
    en_value = clean_card_text(en_value)
    pt_value = clean_card_text(pt_value)
    if pt_value and normalize_text(pt_value) != normalize_text(en_value):
        return f"{en_value} / {pt_value}" if en_value else pt_value
    return en_value


def bilingual_name_parts(value: str) -> set[str]:
    value = clean_card_text(value)
    parts = [part.strip() for part in value.split(" / ") if part.strip()]
    normalized = {normalize_text(part) for part in parts}
    if len(parts) > 1:
        normalized.add(normalize_text(value))
    return {part for part in normalized if part}


def excel_names_match(cell_value: str, search_name: str) -> bool:
    cell_parts = bilingual_name_parts(cell_value)
    search_parts = bilingual_name_parts(search_name)
    if not cell_parts or not search_parts:
        return False
    return bool(cell_parts & search_parts)


def find_name_row_in_excel(name: str) -> int | None:
    if not EXCEL_PATH.exists():
        return None
    workbook = load_workbook(EXCEL_PATH, read_only=True, data_only=True)
    try:
        sheet = workbook.active
        for row_index, row in enumerate(
            sheet.iter_rows(min_col=1, max_col=1, values_only=True),
            start=1,
        ):
            cell_value = row[0] if row else None
            if row_index == 1 and cell_value and str(cell_value).strip() == "Nome":
                continue
            if cell_value is None or str(cell_value).strip() == "":
                continue
            if excel_names_match(str(cell_value), name):
                return row_index
    finally:
        workbook.close()
    return None


def ensure_excel_workbook():
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
    return workbook, sheet


def append_row_to_excel(row: list[str]) -> None:
    workbook, sheet = ensure_excel_workbook()
    sheet.append(row)
    workbook.save(EXCEL_PATH)


def update_row_in_excel(row_number: int, row: list[str]) -> None:
    workbook = load_workbook(EXCEL_PATH)
    sheet = workbook.active
    for col_index, value in enumerate(row, start=1):
        sheet.cell(row=row_number, column=col_index, value=value)
    workbook.save(EXCEL_PATH)


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
class OcrMatchContext:
    ocr: OcrResult
    fractions: list[tuple[str, int]]
    tokens: list[str]
    sets_in_text: set[str]
    hint_card: dict | None = None
    namebar_resolved: bool = False
    hint_from_title: bool = False


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
        self.ocr_lang, self.por_available = detect_tesseract_languages()

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
                    self._ocr(name_crop, "--psm 7", bilingual=True),
                    self._ocr(candidate.crop((0, 0, int(width * 0.9), int(height * 0.18))), "--psm 6", bilingual=True),
                ]
            )
            texts.extend(
                [
                    self._ocr(name_crop, "--psm 7", bilingual=True),
                    self._ocr(bottom_crop, "--psm 6", bilingual=True),
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

    def _ocr(self, image: Image.Image, config: str, bilingual: bool = False) -> str:
        prepared = self._prepare(image)
        lang = self.ocr_lang if bilingual and self.por_available else "eng"
        return pytesseract.image_to_string(prepared, lang=lang, config=config)

    def _ocr_footer(self, image: Image.Image) -> str:
        footer = image.convert("L")
        footer = footer.resize((footer.width * 5, footer.height * 5))
        simple_binary = footer.point(lambda pixel: 255 if pixel > 100 else 0)
        contrast_footer = ImageEnhance.Contrast(footer).enhance(2.5)
        binary = contrast_footer.point(lambda pixel: 255 if pixel > 100 else 0)
        inverted_binary = ImageOps.invert(contrast_footer).point(lambda pixel: 255 if pixel > 100 else 0)
        lang = self.ocr_lang if self.por_available else "eng"
        return "\n".join(
            [
                pytesseract.image_to_string(simple_binary, lang=lang, config="--psm 6"),
                pytesseract.image_to_string(binary, lang=lang, config="--psm 6"),
                pytesseract.image_to_string(inverted_binary, lang=lang, config="--psm 6"),
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
            title_words = re.findall(r"\b[A-ZÀ-Ü][a-zà-ü]{2,}\b", candidate)
            value = min(len(words), 5)
            value += len(title_words) * 3
            if len(title_words) >= 2:
                value += 10
            elif not title_words:
                value -= 6
            value -= sum(3 for word in words if word in ignored)
            value -= sum(2 for word in words if word in {"instant", "sorcery", "artifact", "enchantment", "land"})
            value -= sum(1 for word in words if word.isdigit())
            if re.search(r"[©™]|wizards|coast|wasatch|anthony|palumbo|paliso", candidate, re.IGNORECASE):
                value -= 50
            if "/" in candidate or re.search(r"\b\d{3,}\b", candidate):
                value -= 25
            if len(words) > 6:
                value -= 15
            if re.search(r"\b(criatura|creature|instantanea|feiticaria|encantamento|terreno|artifact)\b", normalized):
                value -= 30
            if re.search(r"\b(da|de|do|das|dos)\b", normalized) and title_words:
                value += 8
            if re.search(r"\b(was|were|nothing|serving|like the|made it|whenever a)\b", normalized):
                value -= 45
            if re.search(r"\b(the|and|for|with|that|this|from|into)\b", normalized) and len(words) >= 5:
                value -= 20
            if "," in candidate:
                value += 12
            if re.match(r"^[A-ZÀ-Ü][a-zà-ü]+,", candidate):
                value += 15
            if re.search(r"[^a-zA-ZÀ-ü0-9\s\-',]", candidate):
                value -= 10
            return value

        line = max(lines, key=score)
        line = re.sub(r"\s+[\dXWUBRGC/]+$", "", line, flags=re.IGNORECASE)
        return clean_ocr_line(line)

    @staticmethod
    def _parse_collector(text: str) -> tuple[str, str, str]:
        text = text.upper().replace("•", " ")
        text = re.sub(r"\bO(\d{2,4})\b", r"0\1", text)
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

        loose_footer = re.search(
            r"\b([CMUR])\s+0*(\d{1,4}[A-Z]?)\b[\s\S]{0,40}?\b([A-Z]{3})\b[\s\S]{0,20}?\b(?:EN|PT)\b",
            text,
        )
        if loose_footer:
            rarity, number, set_code = loose_footer.groups()
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
        self.face_name_choices: list[tuple[str, int]] = []
        self.set_booster_max: dict[str, int] = {}
        self.oracle_collectors: dict[str, set[str]] = {}
        self.loaded = False
        self._norm_cache: dict[str, dict[str, tuple[str, int]]] = {}

    def _normalized_choices(self, key: str, choices: list[tuple[str, int]]) -> dict[str, tuple[str, int]]:
        """Map normalized name -> (original name, card index), built once.

        Fuzzy matching against accent-free keys keeps OCR noise (missing
        accents, mojibake) from steering ``extractOne`` toward the wrong card.
        """
        cache = self.__dict__.setdefault("_norm_cache", {})
        result = cache.get(key)
        if result is None:
            result = {}
            for name, index in choices:
                normalized = normalize_text(name)
                if normalized:
                    result.setdefault(normalized, (name, index))
            cache[key] = result
        return result

    def _pt_norm_choices(self) -> dict[str, tuple[str, int]]:
        return self._normalized_choices("pt", self.pt_name_choices)

    def _en_norm_choices(self) -> dict[str, tuple[str, int]]:
        return self._normalized_choices("en", self.name_choices)

    def _face_norm_choices(self) -> dict[str, tuple[str, int]]:
        return self._normalized_choices("face", self.face_name_choices)

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
        self._norm_cache = {}
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
        self.face_name_choices.clear()
        self.set_booster_max.clear()
        self.oracle_collectors.clear()

        seen_choices = set()
        seen_face_choices = set()
        for index, card in enumerate(self.cards):
            set_code = (card.get("set") or "").upper()
            collector = collector_key(card.get("collector_number"))
            if set_code and collector:
                self.by_print.setdefault((set_code, collector), []).append(card)
            oracle_id = card.get("oracle_id") or card.get("name")
            if oracle_id and collector:
                self.oracle_collectors.setdefault(oracle_id, set()).add(collector)
            if card.get("lang") == "en" and card.get("booster") and str(card.get("collector_number", "")).isdigit():
                collector_number = int(card["collector_number"])
                self.set_booster_max[set_code] = max(self.set_booster_max.get(set_code, 0), collector_number)

            oracle_id = card.get("oracle_id") or card.get("name")
            if card.get("lang") == "pt":
                self.pt_by_oracle.setdefault(oracle_id, []).append(card)
            elif card.get("lang") == "en":
                self.en_by_oracle.setdefault(oracle_id, []).append(card)
                # Art-series prints, tokens and emblems reuse a real card's name
                # (art series often duplicates it as "X // X", and there are token
                # "Goblin"/"Soldier" cards) but are not collectible cards anyone
                # scans to catalog. Keeping them out of the fuzzy name/face pools
                # stops the matcher from returning the art print or token instead
                # of the real card (e.g. AMSH art card instead of the MSC card, or
                # a "Goblin" token instead of "Zhur-Taa Goblin").
                if card.get("layout") in _NON_CATALOG_LAYOUTS:
                    continue
                for candidate in [card.get("name"), card.get("printed_name")]:
                    key = normalize_text(candidate or "")
                    if key and key not in seen_choices:
                        self.name_choices.append((candidate, index))
                        seen_choices.add(key)
                for face in card.get("card_faces") or []:
                    face_name = face.get("name") or ""
                    key = normalize_text(face_name)
                    if key and key not in seen_face_choices:
                        self.face_name_choices.append((face_name, index))
                        seen_face_choices.add(key)

        if MTGJSON_ATOMIC_PATH.exists():
            log("Lendo traduções em português...")
            self.mtgjson_pt_by_name = self._load_mtgjson_portuguese()
            seen_pt_choices = set()
            for index, card in enumerate(self.cards):
                if card.get("lang") != "en":
                    continue
                if card.get("layout") in _NON_CATALOG_LAYOUTS:
                    continue
                pt_card = self.mtgjson_pt_by_name.get(normalize_text(card.get("name", "")))
                pt_name = (pt_card or {}).get("printed_name", "")
                key = normalize_text(pt_name)
                if key and key not in seen_pt_choices:
                    self.pt_name_choices.append((pt_name, index))
                    seen_pt_choices.add(key)
                    for part in re.split(r"\s*/\s*", pt_name):
                        part = part.strip()
                        part_key = normalize_text(part)
                        if part_key and part_key not in seen_face_choices:
                            self.face_name_choices.append((part, index))
                            seen_face_choices.add(part_key)

        state = {
            "cards": self.cards,
            "sets": self.sets,
            "by_print": self.by_print,
            "pt_by_oracle": self.pt_by_oracle,
            "en_by_oracle": self.en_by_oracle,
            "mtgjson_pt_by_name": self.mtgjson_pt_by_name,
            "name_choices": self.name_choices,
            "pt_name_choices": self.pt_name_choices,
            "face_name_choices": self.face_name_choices,
            "set_booster_max": self.set_booster_max,
            "oracle_collectors": self.oracle_collectors,
            "loaded": True,
            "index_version": INDEX_VERSION,
        }
        with INDEX_PATH.open("wb") as handle:
            pickle.dump(state, handle, protocol=pickle.HIGHEST_PROTOCOL)
        self.loaded = True

    OCR_SET_FIXES = {
        "LAIS": "M15",
        "LAAIS": "M15",
        "LAA1S": "M15",
        "LAAI5": "M15",
        "L4A15": "M15",
        "M1S": "M15",
        "M1L5": "M15",
        "FRE": "FRF",
        "ECT": "ECL",
        "ECI": "ECL",
        "EC1": "ECL",
        "RN4": "RNA",
        "RMA": "RNA",
        "RNI": "RNA",
        "AFI": "AFR",
        "AFP": "AFR",
        "GAN": "GTC",
        "GRMN": "GTC",
        "GRAN": "GTC",
        "GRMN": "GTC",
        "ORM": "GTC",
        "FOE": "EOE",
        "EOF": "EOE",
    }

    def _known_set_codes(self) -> set[str]:
        return {set_code for set_code, _collector in self.by_print.keys()}

    def _sets_for_card_count(self, total: int) -> list[str]:
        matches = []
        seen = set()
        for code, info in self.sets.items():
            if info.get("card_count") == total and code not in seen:
                matches.append(code)
                seen.add(code)
        for code, booster_max in self.set_booster_max.items():
            if booster_max == total and code not in seen:
                matches.append(code)
                seen.add(code)
        return matches

    @staticmethod
    def _is_likely_collector_fraction(collector: str, total: int) -> bool:
        if total < 30:
            return False
        collector_digits = re.sub(r"\D", "", collector)
        if not collector_digits:
            return False
        return int(collector_digits) <= total

    @staticmethod
    def _collector_fractions_in_text(raw_text: str) -> list[tuple[str, int]]:
        fractions = []
        text = raw_text.upper()
        for match in re.finditer(r"\b([0-9OIL]{1,4}[A-Z]?)\s*/\s*([0-9OILSE]{1,5})\b", text):
            collector = ocr_collector_key(match.group(1))
            total_raw = match.group(2).translate(str.maketrans({"O": "0", "I": "1", "L": "1", "E": "8", "S": "5"}))
            if re.fullmatch(r"\d{1,5}", total_raw):
                total = int(total_raw)
                if ScryfallDatabase._is_likely_collector_fraction(collector, total):
                    fractions.append((collector, total))
        return fractions

    def _corrected_set(
        self,
        token: str,
        known_sets: set[str],
        collector: str = "",
        total: int = 0,
    ) -> str:
        if token in known_sets:
            return token
        fixed = self.OCR_SET_FIXES.get(token, "")
        if fixed and fixed in known_sets:
            return fixed
        if total and collector:
            matching_sets = self._sets_for_card_count(total)
            if len(matching_sets) == 1 and self.by_print.get((matching_sets[0], collector)):
                return matching_sets[0]
        if fuzz and process and len(token) >= 2:
            candidates = list(known_sets)
            if total and collector:
                count_sets = self._sets_for_card_count(total)
                if count_sets:
                    candidates = count_sets
            match = process.extractOne(token, candidates, scorer=fuzz.ratio)
            min_score = 86 if len(token) <= 3 else 80
            if match and match[1] >= min_score:
                if self.by_print.get((match[0], collector)):
                    return match[0]
        return ""

    def _print_candidates(self, set_code: str, collector: str) -> list[dict]:
        candidates = [card for card in self.by_print.get((set_code, collector), []) if card.get("lang") == "en"]
        return candidates or self.by_print.get((set_code, collector), [])

    def _pick_best_print_candidate(self, candidates: list[dict], raw_text: str) -> dict | None:
        if not candidates:
            return None
        if len(candidates) == 1:
            return candidates[0]
        scored = []
        for card in candidates:
            score = 0
            if self._card_name_in_raw_text(card, raw_text):
                score += 100
            pt_card = self.portuguese_for(card)
            pt_name = (pt_card or {}).get("printed_name") or ""
            if pt_name and normalize_text(pt_name) in normalize_text(raw_text):
                score += 80
            scored.append((score, card))
        scored.sort(key=lambda item: item[0], reverse=True)
        if scored[0][0] > 0:
            return scored[0][1]
        return candidates[0]

    def _set_codes_in_text(self, raw_text: str) -> set[str]:
        known_sets = self._known_set_codes()
        tokens = re.findall(r"[A-Z0-9]+", raw_text.upper())
        found = set()
        for index, token in enumerate(tokens):
            if token in known_sets and self._set_token_has_print_context(tokens, index):
                found.add(token)
                continue
            fixed = self.OCR_SET_FIXES.get(token, "")
            if fixed and fixed in known_sets and self._set_fix_has_print_context(token, tokens, index):
                found.add(fixed)
        if fuzz and process:
            for index, token in enumerate(tokens):
                if len(token) != 3 or token in found or not self._set_token_has_print_context(tokens, index):
                    continue
                match = process.extractOne(token, list(known_sets), scorer=fuzz.ratio)
                if match and match[1] >= 88:
                    found.add(match[0])
        found.update(self._symbol_set_codes_in_text(raw_text, known_sets))
        return found

    @staticmethod
    def _set_token_has_print_context(tokens: list[str], index: int) -> bool:
        token = tokens[index]
        left_window = tokens[max(0, index - 4) : index]
        right_window = tokens[index + 1 : index + 5]
        if token not in AMBIGUOUS_SET_CODE_WORDS and not (
            len(token) == 3 and token.isalpha()
        ):
            return True
        if any(value in {"EN", "PT", "ES", "FR", "DE", "IT", "JP", "KO"} for value in right_window):
            return True
        for value in left_window:
            number_match = re.fullmatch(r"0*(\d{3,4})[A-Z]?", value)
            if number_match and int(number_match.group(1)) <= 999:
                return True
        return False

    @classmethod
    def _set_fix_has_print_context(cls, token: str, tokens: list[str], index: int) -> bool:
        if not (len(token) == 3 and token.isalpha()):
            return True
        return cls._set_token_has_print_context(tokens, index)

    @staticmethod
    def _symbol_set_codes_in_text(raw_text: str, known_sets: set[str]) -> set[str]:
        found = set()
        for line in raw_text.splitlines():
            normalized = normalize_text(line)
            if not re.search(
                r"\b(criatura|creature|instantanea|instant|feiticaria|sorcery|"
                r"encantamento|enchantment|artefato|artifact|terreno|land)\b",
                normalized,
            ):
                continue
            tokens = re.findall(r"[A-Z0-9]+", line.upper())
            if tokens and tokens[-1] in {"X", "XX"} and "10E" in known_sets:
                found.add("10E")
        return found

    def _namebar_candidates(self, namebar_text: str) -> list[str]:
        candidates = []
        for line in ranked_namebar_lines(namebar_text):
            for value in (line, normalize_ocr_name(line)):
                if value and value not in candidates:
                    candidates.append(value)
        normalized_all = normalize_ocr_name(namebar_text)
        if normalized_all and normalized_all not in candidates and is_plausible_namebar(normalized_all):
            candidates.append(normalized_all)
        return candidates

    def _best_name_coverage(self, card: dict, line: str) -> int:
        """Best coverage of any of the card's names against a namebar line."""
        normalized_line = normalize_text(line)
        names = [card.get("name", "")]
        for face in card.get("card_faces") or []:
            if face.get("name"):
                names.append(face["name"])
        pt_card = self.portuguese_for(card)
        if pt_card:
            printed = pt_card.get("printed_name") or pt_card.get("name") or ""
            names.append(printed)
            names.extend(part.strip() for part in re.split(r"\s*/\s*", printed))
        return max((self._hint_word_coverage(name, normalized_line) for name in names if name), default=0)

    def _fuzzy_namebar(self, namebar_text: str) -> dict | None:
        # Evaluate every plausible namebar line and keep the match whose name is
        # most fully present in the line it came from. This stops a short name
        # (e.g. the plane "Naya") from winning over a fuller one ("Medalhão de
        # Naya") just because a noisier line was scored higher.
        best_card = None
        best_cov = 0
        for candidate in self._namebar_candidates(namebar_text):
            card = self._fuzzy_name(candidate, strict_words=True, min_score=84, prefer_en=True)
            if not card and is_plausible_namebar(candidate):
                card = self._fuzzy_name(candidate, strict_words=False, min_score=76, prefer_en=True)
            if not card:
                continue
            cov = self._best_name_coverage(card, candidate)
            if cov > best_cov:
                best_cov = cov
                best_card = card
        return best_card

    def _find_by_face_names_in_text(self, raw_text: str) -> tuple[dict | None, str]:
        if not raw_text or not self.face_name_choices or not process or not fuzz:
            return None, ""
        choices = self._face_norm_choices()
        best_card = None
        best_name = ""
        best_score = 0
        for line in raw_text.splitlines():
            cleaned = clean_ocr_line(line)
            normalized = normalize_text(cleaned)
            words = normalized.split()
            if not words or len(words) > 4:
                continue
            if not any(len(word) >= 4 for word in words):
                continue
            for size in range(min(3, len(words)), 0, -1):
                for start in range(0, len(words) - size + 1):
                    candidate = " ".join(words[start : start + size])
                    if len(candidate) < 5:
                        continue
                    if candidate in COMMON_EN_OCR_WORDS or candidate in COMMON_PT_OCR_WORDS:
                        continue
                    match = process.extractOne(candidate, choices.keys(), scorer=fuzz.WRatio)
                    if not match:
                        continue
                    original_name, card_index = choices[match[0]]
                    matched_len = len(match[0])
                    min_score = 96 if len(candidate) < 7 else 90
                    if match[1] < min_score:
                        continue
                    if len(candidate) < matched_len * 0.55:
                        continue
                    matched_words = match[0].split()
                    if len(matched_words) == 1 and match[0] not in candidate.split():
                        continue
                    if match[1] > best_score:
                        best_score = match[1]
                        best_name = original_name
                        best_card = self.cards[card_index]
        if best_card:
            return best_card, best_name
        return None, ""

    def _fraction_supports_print(self, set_code: str, collector: str, fractions: list[tuple[str, int]]) -> int:
        score = 0
        for fraction_collector, total in fractions:
            if fraction_collector != collector:
                continue
            if set_code in self._sets_for_card_count(total):
                score = max(score, 150)
        return score

    def _distinctive_name_in_raw_text(self, card: dict, raw_text: str) -> bool:
        names = [card.get("name", "")]
        for face in card.get("card_faces") or []:
            face_name = face.get("name") or ""
            if face_name:
                names.append(face_name)
        pt_card = self.portuguese_for(card)
        if pt_card:
            printed_name = pt_card.get("printed_name") or pt_card.get("name") or ""
            names.append(printed_name)
            for part in re.split(r"\s*/\s*", printed_name):
                part = part.strip()
                if part:
                    names.append(part)
        normalized_raw = normalize_text(raw_text)
        raw_words = [word for word in normalized_raw.split() if len(word) >= 3]
        for card_name in names:
            if not card_name:
                continue
            normalized_name = normalize_text(card_name)
            if normalized_name and normalized_name in normalized_raw:
                return True
            name_words = [
                word for word in normalized_name.split() if len(word) >= 4 and word not in COMMON_PT_OCR_WORDS
            ]
            if not name_words:
                name_words = [word for word in normalized_name.split() if len(word) >= 4]
            if not name_words:
                continue
            if all(word in COMMON_PT_OCR_WORDS for word in name_words):
                continue
            required = 2 if len(name_words) >= 2 else 1
            if fuzzy_word_count(name_words, raw_words) >= required:
                return True
            if fuzz:
                for line in raw_text.splitlines():
                    cleaned = clean_ocr_line(line)
                    if len(cleaned.split()) > 6:
                        continue
                    normalized_line = normalize_ocr_name(cleaned)
                    line_words = [word for word in normalized_line.split() if len(word) >= 3]
                    if name_words and len(name_words) == 1:
                        word = name_words[0]
                        if not any(
                            words_fuzzy_match(word, raw_word, 88) and len(raw_word) >= max(5, len(word) * 0.55)
                            for raw_word in line_words
                        ):
                            continue
                    if fuzz.WRatio(card_name, cleaned) >= 84 or fuzz.WRatio(card_name, normalized_line) >= 76:
                        return True
        return False

    def _print_evidence_in_ocr(self, card: dict, ctx: OcrMatchContext) -> bool:
        ocr = ctx.ocr
        set_code = (card.get("set") or "").upper()
        collector = collector_key(card.get("collector_number"))
        if ocr.set_code and ocr.collector_number:
            if ocr.set_code.upper() == set_code and collector_key(ocr.collector_number) == collector:
                return True
        for index, token in enumerate(ctx.tokens):
            if collector_key(token) != collector:
                continue
            window = set(ctx.tokens[index + 1 : index + 16])
            if set_code in window:
                return True
        fraction_collectors = {fraction_collector for fraction_collector, _total in ctx.fractions}
        if set_code in ctx.sets_in_text and collector in fraction_collectors:
            return True
        if collector in self._collector_numbers_in_text(ocr.raw_text):
            if set_code in ctx.sets_in_text or set_code in ctx.tokens:
                return True
        return False

    def _reliable_hint_card(self, ocr: OcrResult) -> dict | None:
        # Prefer the name bar (the card title) over face names: a single common
        # word in the body (e.g. the creature type "Goblin") can match a split
        # card's face ("Weird // Goblin") and give a misleading hint.
        namebar_en = self._fuzzy_namebar(ocr.namebar_text)
        if namebar_en and self._distinctive_name_in_raw_text(namebar_en, ocr.raw_text):
            return namebar_en
        namebar_pt, _ = self._fuzzy_portuguese_text(
            best_namebar_line(ocr.namebar_text) or ocr.namebar_text,
            namebar_mode=True,
        )
        if namebar_pt and self._distinctive_name_in_raw_text(namebar_pt, ocr.raw_text):
            return namebar_pt
        face_card, _ = self._find_by_face_names_in_text(ocr.raw_text)
        if face_card and self._distinctive_name_in_raw_text(face_card, ocr.raw_text):
            return face_card
        pt_words_card, _ = self._fuzzy_portuguese_words_in_text(ocr.raw_text)
        if pt_words_card and self._distinctive_name_in_raw_text(pt_words_card, ocr.raw_text):
            return pt_words_card
        if ocr.name_hint and self._likely_name_hint(ocr.name_hint):
            normalized_hint = normalize_ocr_name(ocr.name_hint)
            if is_plausible_namebar(normalized_hint):
                card = self._fuzzy_name(normalized_hint, strict_words=False, min_score=76)
                if card and self._distinctive_name_in_raw_text(card, ocr.raw_text):
                    return card
            if is_plausible_namebar(ocr.name_hint):
                card = self._fuzzy_name(ocr.name_hint, strict_words=True, min_score=84)
                if card and self._distinctive_name_in_raw_text(card, ocr.raw_text):
                    return card
        return None

    def _fuzzy_portuguese_words_in_text(self, raw_text: str) -> tuple[dict | None, str]:
        if not raw_text or not self.pt_name_choices:
            return None, ""
        raw_words = [word for word in normalize_text(raw_text).split() if len(word) >= 4]
        if len(raw_words) < 2:
            return None, ""
        best_word_card = None
        best_word_name = ""
        best_word_score = 0
        for pt_name, index in self.pt_name_choices:
            pt_words = [word for word in normalize_text(pt_name).split() if len(word) >= 4]
            if not pt_words:
                continue
            distinctive_words = [word for word in pt_words if word not in COMMON_PT_OCR_WORDS]
            match_words = distinctive_words or pt_words
            all_pt_words = [word for word in normalize_text(pt_name).split() if len(word) >= 3]
            first_word = all_pt_words[0] if all_pt_words else ""
            if len(all_pt_words) > 1 and first_word:
                if not any(words_fuzzy_match(first_word, word, 88) for word in raw_words):
                    continue
            if len(match_words) == 1 and match_words[0] not in raw_words:
                continue
            matches = fuzzy_word_count(match_words, raw_words)
            required = min(3, len(match_words)) if len(match_words) >= 3 else min(2, len(match_words))
            if matches < required:
                continue
            score = matches * 100 + min(len(pt_words), 6)
            if score > best_word_score:
                best_word_score = score
                best_word_name = pt_name
                best_word_card = self.cards[index]
        if best_word_card:
            return best_word_card, best_word_name
        return None, ""

    def _oracle_has_collector(self, oracle_id: str, collector: str) -> bool:
        return collector in self.oracle_collectors.get(oracle_id, set())

    def _build_match_context(self, ocr: OcrResult) -> OcrMatchContext:
        fractions = self._collector_fractions_in_text(ocr.raw_text)
        tokens = re.findall(r"[A-Z0-9]+", ocr.raw_text.upper())
        sets_in_text = self._set_codes_in_text(ocr.raw_text)
        namebar_en = self._fuzzy_namebar(ocr.namebar_text)
        namebar_pt, _ = self._fuzzy_portuguese_text(
            best_namebar_line(ocr.namebar_text) or ocr.namebar_text,
            namebar_mode=True,
        )
        face_card, _ = self._find_by_face_names_in_text(ocr.raw_text)
        pt_words_card, _ = self._fuzzy_portuguese_words_in_text(ocr.raw_text)
        hint_card = self._reliable_hint_card(ocr)
        hint_oracle = hint_card.get("oracle_id") if hint_card else None
        hint_from_title = bool(
            hint_oracle
            and (
                (namebar_en and namebar_en.get("oracle_id") == hint_oracle)
                or (namebar_pt and namebar_pt.get("oracle_id") == hint_oracle)
            )
        )
        return OcrMatchContext(
            ocr=ocr,
            fractions=fractions,
            tokens=tokens,
            sets_in_text=sets_in_text,
            hint_card=hint_card,
            namebar_resolved=bool(
                (namebar_en and self._card_name_in_raw_text(namebar_en, ocr.raw_text))
                or (namebar_pt and self._card_name_in_raw_text(namebar_pt, ocr.raw_text))
                or (face_card and self._card_name_in_raw_text(face_card, ocr.raw_text))
                or pt_words_card
            ),
            hint_from_title=hint_from_title,
        )

    def _find_by_collector_fraction(self, raw_text: str, ctx: OcrMatchContext | None = None) -> tuple[dict | None, str, str]:
        fractions = ctx.fractions if ctx else self._collector_fractions_in_text(raw_text)
        if not fractions:
            return None, "", ""
        sets_in_text = ctx.sets_in_text if ctx else self._set_codes_in_text(raw_text)
        collectors_in_fractions = {collector for collector, _total in fractions}
        if sets_in_text and collectors_in_fractions:
            best_card = None
            best_set = ""
            best_collector = ""
            best_score = -1
            for set_code in sets_in_text:
                for collector in collectors_in_fractions:
                    candidates = self._print_candidates(set_code, collector)
                    if not candidates:
                        continue
                    card = self._pick_best_print_candidate(candidates, raw_text)
                    if not card:
                        continue
                    score = 250
                    score += self._fraction_supports_print(set_code, collector, fractions)
                    if self._card_name_in_raw_text(card, raw_text):
                        score += 100
                    if score > best_score:
                        best_score = score
                        best_card = card
                        best_set = set_code
                        best_collector = collector
            if best_card:
                return best_card, best_set, best_collector

        fraction_counts: dict[tuple[str, int], int] = {}
        for collector, total in fractions:
            fraction_counts[(collector, total)] = fraction_counts.get((collector, total), 0) + 1
        ranked_fractions = sorted(fraction_counts.items(), key=lambda item: item[1], reverse=True)
        for (collector, total), _count in ranked_fractions:
            matching_sets = self._sets_for_card_count(total)
            if not matching_sets:
                continue
            best_card = None
            best_set = ""
            best_score = -1
            for set_code in matching_sets:
                candidates = self._print_candidates(set_code, collector)
                if not candidates:
                    continue
                card = self._pick_best_print_candidate(candidates, raw_text)
                if not card:
                    continue
                score = 50
                score += self._fraction_supports_print(set_code, collector, fractions)
                if set_code in sets_in_text:
                    score += 200
                if self._card_name_in_raw_text(card, raw_text):
                    score += 100
                if score > best_score:
                    best_score = score
                    best_card = card
                    best_set = set_code
            if best_card:
                if len(matching_sets) > 1 and best_score < 150:
                    continue
                return best_card, best_set, collector
        return None, "", ""

    def _find_by_fraction_and_name(
        self, raw_text: str, ctx: OcrMatchContext
    ) -> tuple[dict | None, str, str]:
        """Strongest possible signal: a collector fraction (e.g. "198/249")
        whose printing also has its name in the OCR text. The footer number and
        the card name agree, so this beats every fuzzy path — and it is robust
        to a mangled set code (``IMA`` read as ``IM AS``) because the set is
        derived from the card-count total instead of the OCR'd letters."""
        for collector, total in ctx.fractions:
            candidate_sets = set(ctx.sets_in_text) | set(self._sets_for_card_count(total))
            for set_code in candidate_sets:
                for card in self._print_candidates(set_code, collector):
                    if self._strong_card_name_in_raw_text(card, raw_text):
                        return card, set_code, collector
        return None, "", ""

    def find(self, ocr: OcrResult) -> tuple[dict, str]:
        self.load()
        ctx = self._build_match_context(ocr)

        def finalize(card: dict, reason: str) -> tuple[dict, str]:
            return self._prefer_print_from_ocr(card, ocr.raw_text), reason

        def try_match(card: dict | None, reason: str) -> tuple[dict, str] | None:
            if not card:
                return None
            card = self._prefer_print_from_ocr(card, ocr.raw_text)
            if self.is_confident_match(card, ctx):
                return card, reason
            return None

        if ocr.set_code and ocr.collector_number:
            candidates = self._print_candidates(ocr.set_code.upper(), collector_key(ocr.collector_number))
            if candidates:
                card = self._pick_best_print_candidate(candidates, ocr.raw_text)
                matched = try_match(card, f"código {ocr.set_code.upper()} #{ocr.collector_number}")
                if matched:
                    return matched

        card, set_code, collector_number = self._find_by_fraction_and_name(ocr.raw_text, ctx)
        if card:
            return finalize(card, f"fração+nome {set_code} #{collector_number}")

        card, set_code, collector_number = self._find_print_in_ocr_text(ocr.raw_text)
        matched = try_match(card, f"rodapé OCR {set_code} #{collector_number}")
        if matched:
            return matched

        card, matched_line = self._exact_english_title_in_text(ocr.raw_text)
        matched = try_match(card, f"tÃ­tulo OCR '{matched_line}'")
        if matched:
            return matched

        face_card, face_name = self._find_by_face_names_in_text(ocr.raw_text)
        if face_card and self._distinctive_name_in_raw_text(face_card, ocr.raw_text):
            matched = try_match(face_card, f"face OCR '{face_name}'")
            if matched:
                return matched

        pt_words_card, pt_words_name = self._fuzzy_portuguese_words_in_text(ocr.raw_text)
        if pt_words_card and self._distinctive_name_in_raw_text(pt_words_card, ocr.raw_text):
            matched = try_match(pt_words_card, f"nome PT palavras '{pt_words_name}'")
            if matched:
                return matched

        namebar_en = self._fuzzy_namebar(ocr.namebar_text)
        namebar_line = best_namebar_line(ocr.namebar_text) or normalize_ocr_name(ocr.namebar_text)
        if namebar_en and (
            self._distinctive_name_in_raw_text(namebar_en, ocr.raw_text)
            or self._card_name_matches(namebar_en, namebar_line)
        ):
            matched = try_match(namebar_en, "nome superior OCR")
            if matched:
                return matched

        namebar_pt, namebar_pt_name = self._fuzzy_portuguese_text(
            best_namebar_line(ocr.namebar_text) or ocr.namebar_text,
            namebar_mode=True,
        )
        if namebar_pt and self._card_name_in_raw_text(namebar_pt, ocr.raw_text):
            matched = try_match(namebar_pt, f"nome superior PT OCR '{namebar_pt_name}'")
            if matched:
                return matched

        card, set_code, collector_number = self._find_by_collector_fraction(ocr.raw_text, ctx)
        matched = try_match(card, f"fração collector {set_code} #{collector_number}")
        if matched:
            return matched

        if not ctx.namebar_resolved:
            card, matched_name = self._fuzzy_portuguese_text(ocr.raw_text)
            matched = try_match(card, f"nome PT OCR '{matched_name}'")
            if matched:
                return matched

        if ocr.name_hint and self._likely_name_hint(ocr.name_hint):
            normalized_hint = normalize_ocr_name(ocr.name_hint)
            card = None
            if is_plausible_namebar(normalized_hint):
                card = self._fuzzy_name(normalized_hint, strict_words=False, min_score=76)
            if not card and is_plausible_namebar(ocr.name_hint):
                card = self._fuzzy_name(ocr.name_hint, strict_words=True, min_score=84)
            matched = try_match(card, f"nome OCR '{ocr.name_hint}'")
            if matched:
                return matched

        card, matched_line = self._fuzzy_ocr_text(ocr.raw_text)
        matched = try_match(card, f"texto OCR '{matched_line}'")
        if matched:
            return matched
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

    def _print_from_context(self, same_oracle_cards: list[dict], raw_text: str) -> dict | None:
        """Pick a printing from footer/copyright evidence other than an exact
        collector match: the set card-count total (e.g. ".../81" -> the only
        printing in an 81-card set) and then the copyright year (e.g. the
        "© 1993-2011" range disambiguates the 2011 printing)."""
        if not same_oracle_cards:
            return None
        totals = {total for _collector, total in self._collector_fractions_in_text(raw_text)}
        for total in totals:
            count_sets = set(self._sets_for_card_count(total))
            matches = [item for item in same_oracle_cards if (item.get("set") or "").upper() in count_sets]
            if len(matches) == 1:
                return matches[0]
        years = re.findall(r"\b(19\d{2}|20\d{2})\b", raw_text)
        if years:
            target = max(years)
            matches = [item for item in same_oracle_cards if (item.get("released_at") or "")[:4] == target]
            if len(matches) == 1:
                return matches[0]
        # The copyright year is often the only thing separating reprints, and it
        # sits on the lowest-contrast line, so OCR mangles a digit (e.g. "2010"
        # read as "1010"). Recover it by matching each printing's release year
        # against the OCR allowing one wrong digit.
        release_years = {(item.get("released_at") or "")[:4] for item in same_oracle_cards}
        release_years = {year for year in release_years if re.fullmatch(r"\d{4}", year)}
        tolerant = self._years_in_copyright(raw_text, release_years)
        if tolerant:
            target = max(tolerant)
            matches = [item for item in same_oracle_cards if (item.get("released_at") or "")[:4] == target]
            if len(matches) == 1:
                return matches[0]
        return None

    def _flavor_scores(self, cards: list[dict], raw_text: str) -> dict[int, tuple[int, int]]:
        """(proper-noun hits, content-word hits) of each card's flavor text in
        the OCR. Flavor is stored in English, but proper nouns in an attribution
        ("—Sheoldred, Whispering One") are spelled the same across languages, so
        they still anchor a match against a Portuguese scan."""
        raw_words = set(normalize_text(raw_text).split())
        scores: dict[int, tuple[int, int]] = {}
        for item in cards:
            flavor = card_faces_text(item, "flavor_text")
            if not flavor:
                scores[id(item)] = (0, 0)
                continue
            proper = {
                normalize_text(word) for word in re.findall(r"[A-Z][A-Za-z'-]{3,}", flavor)
            }
            proper = {word for word in proper if word}
            proper_hits = 0
            for token in proper:
                if token in raw_words:
                    proper_hits += 1
                elif fuzz and any(fuzz.ratio(token, word) >= 85 for word in raw_words):
                    proper_hits += 1
            content_words = {word for word in normalize_text(flavor).split() if len(word) >= 4}
            content_hits = sum(1 for word in content_words if word in raw_words)
            scores[id(item)] = (proper_hits, content_hits)
        return scores

    def _flavor_compatible_prints(self, cards: list[dict], raw_text: str) -> list[dict]:
        """Drop printings whose flavor text contradicts what the OCR read. Many
        reprints differ only by flavor (Trained Armodon's Tempest quote is "These
        are its last days...", but the reprints say "Armodons are trained to step
        on things"), so a printing whose flavor is absent from the scan cannot be
        the card in hand. Only filters when at least one printing's flavor is
        strongly present, so a flavorless or badly garbled scan is left intact."""
        if len(cards) <= 1:
            return cards
        scores = self._flavor_scores(cards, raw_text)
        best = max(scores.values())
        if best[0] < 1 and best[1] < 3:
            return cards
        keep = [item for item in cards if scores[id(item)] == best]
        return keep or cards

    @staticmethod
    def _years_in_copyright(raw_text: str, candidate_years: set[str]) -> set[str]:
        normalized = re.sub(r"[OoQ]", "0", raw_text)
        normalized = re.sub(r"[lI|]", "1", normalized)
        tokens = set(re.findall(r"\d{4}", normalized))
        found = set()
        for year in candidate_years:
            for token in tokens:
                if sum(a != b for a, b in zip(year, token)) <= 1:
                    found.add(year)
                    break
        return found

    @staticmethod
    def _earliest_print(same_oracle_cards: list[dict], card: dict) -> dict:
        if not same_oracle_cards:
            return card
        return sorted(same_oracle_cards, key=lambda item: item.get("released_at") or "9999-99-99")[0]

    def _prefer_print_from_ocr(self, card: dict, raw_text: str) -> dict:
        oracle_id = card.get("oracle_id")
        set_codes_in_text = self._set_codes_in_text(raw_text)
        same_oracle_cards = [
            item
            for item in self.cards
            if item.get("lang") == "en" and item.get("oracle_id") == oracle_id
        ]
        collectors = self._collector_numbers_in_text(raw_text)

        def set_in_text(item: dict) -> bool:
            code = (item.get("set") or "").upper()
            return bool(code) and code in set_codes_in_text

        set_matches = [item for item in same_oracle_cards if set_in_text(item)]
        if set_matches:
            # When a collector number was read from the card, keep the printing
            # whose set AND collector both appear in the OCR text. Otherwise an
            # exact set+collector match gets overwritten by an arbitrary same-set
            # sibling (a serialized or alternate-art reprint under a different
            # collector number), which is how MSC #58 turned into MSC #369.
            if collectors:
                for item in set_matches:
                    if collector_key(item.get("collector_number")) in collectors:
                        return item
                if card in set_matches:
                    return card
            return set_matches[0]

        # With no readable set/collector, the flavor text on the card is the
        # strongest remaining signal for which reprint this is: discard the
        # printings whose flavor contradicts the scan before falling back.
        compatible = self._flavor_compatible_prints(same_oracle_cards, raw_text)
        context_pick = self._print_from_context(compatible, raw_text)
        if not collectors:
            return context_pick or self._earliest_print(compatible, card)
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
                if (item.get("set") or "").upper() in set_codes_in_text:
                    return item
            return same_cards[0]
        return context_pick or self._earliest_print(compatible, card)

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
        known_sets = self._known_set_codes()
        fractions = self._collector_fractions_in_text(raw_text)
        fraction_totals = {collector: total for collector, total in fractions}
        language_codes = {"EN", "PT", "ES", "FR", "DE", "IT", "JP", "KO", "RU", "ZHS", "ZHT"}
        rarity_codes = {"C", "U", "R", "M"}
        tokens = re.findall(r"[A-Z0-9]+", raw_text.upper())

        def candidate_from(set_code: str, collector: str) -> dict | None:
            candidates = self._print_candidates(set_code, collector)
            if not candidates:
                return None
            for card in candidates:
                if self._card_name_in_raw_text(card, raw_text):
                    return card
            total = fraction_totals.get(collector, 0)
            collector_digits = re.sub(r"\D", "", collector)
            if collector_digits and int(collector_digits) <= 9 and not total:
                return None
            if total:
                if set_code in self._sets_for_card_count(total):
                    return self._pick_best_print_candidate(candidates, raw_text)
                return None
            if self.by_print.get((set_code, collector)):
                return self._pick_best_print_candidate(candidates, raw_text)
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
            total = fraction_totals.get(collector, 0)
            for lookahead in range(collector_index + 1, min(collector_index + 25, len(tokens) - 1)):
                set_code = self._corrected_set(tokens[lookahead], known_sets, collector, total)
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
            total = fraction_totals.get(collector, 0)
            if index + 1 < len(tokens) and tokens[index + 1].isdigit():
                search_start = index + 2
            else:
                search_start = index + 1
            for lookahead in range(search_start, min(search_start + 25, len(tokens) - 1)):
                set_code = self._corrected_set(tokens[lookahead], known_sets, collector, total)
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
        for face in card.get("card_faces") or []:
            face_name = face.get("name") or ""
            if face_name:
                names.append(face_name)
        pt_card = self.portuguese_for(card)
        if pt_card:
            names.append(pt_card.get("printed_name") or pt_card.get("name") or "")
            printed_name = (pt_card or {}).get("printed_name") or ""
            for part in re.split(r"\s*/\s*", printed_name):
                part = part.strip()
                if part:
                    names.append(part)
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

    def _strong_card_name_in_raw_text(self, card: dict, raw_text: str) -> bool:
        names = [card.get("name", "")]
        for face in card.get("card_faces") or []:
            face_name = face.get("name") or ""
            if face_name:
                names.append(face_name)
        pt_card = self.portuguese_for(card)
        if pt_card:
            printed_name = pt_card.get("printed_name") or pt_card.get("name") or ""
            names.append(printed_name)
            names.extend(part.strip() for part in re.split(r"\s*/\s*", printed_name) if part.strip())

        normalized_raw = normalize_text(raw_text)
        raw_words = [word for word in normalized_raw.split() if len(word) >= 3]
        for card_name in (name for name in names if name):
            normalized_name = normalize_text(card_name)
            if normalized_name and normalized_name in normalized_raw:
                return True
            name_words = [
                word
                for word in normalized_name.split()
                if len(word) >= 3 and word not in COMMON_EN_OCR_WORDS and word not in COMMON_PT_OCR_WORDS
            ]
            if len(name_words) < 2:
                continue
            required = min(3, len(name_words))
            if fuzzy_word_count(name_words, raw_words, threshold=88) >= required:
                return True
        return False

    def _hint_card_shares_set(self, hint_card: dict, set_code: str) -> bool:
        if not set_code or not hint_card:
            return False
        oracle_id = hint_card.get("oracle_id") or hint_card.get("name")
        prints = self.en_by_oracle.get(oracle_id) or self.pt_by_oracle.get(oracle_id) or []
        return any((item.get("set") or "").upper() == set_code for item in prints)

    def is_confident_match(self, card: dict, ctx: OcrMatchContext) -> bool:
        ocr = ctx.ocr
        name_present = self._card_name_in_raw_text(card, ocr.raw_text)
        is_hint_card = bool(ctx.hint_card and ctx.hint_card.get("oracle_id") == card.get("oracle_id"))
        exact_title_card, _ = self._exact_english_title_in_text(ocr.raw_text)
        has_exact_title = bool(
            exact_title_card and exact_title_card.get("oracle_id") == card.get("oracle_id")
        )
        # When this card's own name is the title read clearly from the image, a
        # collector number that carries an OCR digit error (e.g. 215 read as 219)
        # must not veto it — the name is the stronger signal.
        strong_name = (name_present and is_hint_card) or has_exact_title
        if not strong_name and self._ocr_contradicts_card(card, ctx):
            return False
        if (
            ctx.hint_card
            and ctx.hint_card.get("oracle_id") != card.get("oracle_id")
            and not strong_name
        ):
            if not self._print_evidence_in_ocr(card, ctx):
                return False
            if not self._card_name_in_raw_text(card, ocr.raw_text):
                # The card's own name is absent from the OCR, so this candidate
                # rests on a set+collector pair alone. When the title bar clearly
                # read a different card (hint_from_title), that title outweighs a
                # pair recovered from noisy footer/body tokens: trust the pair
                # only if the dedicated footer parser read a structured
                # set_code + collector_number that names this card exactly.
                # Otherwise it is a hallucinated footer match, e.g. the "Golias
                # Zumbi" scan turning into "Dread Reaper POR #89".
                structured_match = bool(
                    ocr.set_code
                    and ocr.collector_number
                    and ocr.set_code.upper() == (card.get("set") or "").upper()
                    and collector_key(ocr.collector_number)
                    == collector_key(card.get("collector_number"))
                )
                if ctx.hint_from_title and not structured_match:
                    return False
                # A name read clearly from the card must outweigh a collector
                # number when the named card also exists in this set: that is the
                # signature of a misread digit (e.g. 215 -> 219 turning
                # "Zhur-Taa Goblin" into "Senate Griffin").
                if self._hint_card_shares_set(ctx.hint_card, (card.get("set") or "").upper()):
                    return False
        if self._card_name_in_raw_text(card, ocr.raw_text):
            return True
        if ctx.hint_card and ctx.hint_card.get("oracle_id") == card.get("oracle_id"):
            return True
        set_code = (card.get("set") or "").upper()
        collector = collector_key(card.get("collector_number"))
        for index, token in enumerate(ctx.tokens):
            if collector_key(token) != collector:
                continue
            window = set(ctx.tokens[index + 1 : index + 16])
            if set_code in window:
                return True
        for fraction_collector, total in ctx.fractions:
            if fraction_collector != collector:
                continue
            matching_sets = self._sets_for_card_count(total)
            if set_code not in matching_sets:
                continue
            if set_code in ctx.sets_in_text:
                return True
            if len(matching_sets) == 1 and set_code in ctx.tokens:
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

    def _ocr_contradicts_card(self, card: dict, ctx: OcrMatchContext) -> bool:
        raw_text = ctx.ocr.raw_text
        fractions = ctx.fractions
        collectors = self._collector_numbers_in_text(raw_text)
        if not fractions and not collectors:
            return False
        card_collector = collector_key(card.get("collector_number"))
        card_set = (card.get("set") or "").upper()
        oracle_id = card.get("oracle_id")

        fraction_collectors = {collector for collector, _total in fractions}
        if card_collector in fraction_collectors:
            if card_set in ctx.sets_in_text:
                return False
            if self._card_name_in_raw_text(card, raw_text):
                return False
            supporting = False
            for collector, total in fractions:
                if collector != card_collector:
                    continue
                if card_set in self._sets_for_card_count(total):
                    supporting = True
                    break
            return not supporting

        if card_collector in collectors and self._oracle_has_collector(oracle_id, card_collector):
            return False

        if fraction_collectors:
            for collector, total in fractions:
                if self._oracle_has_collector(oracle_id, collector):
                    continue
                matching_sets = self._sets_for_card_count(total)
                for set_code in matching_sets:
                    if self.by_print.get((set_code, collector)):
                        return True

        if collectors and not self._oracle_has_collector(oracle_id, card_collector):
            for collector in collectors:
                if self._oracle_has_collector(oracle_id, collector):
                    continue
                if collector != card_collector:
                    return True
        return False

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
            "serving",
            "nothing",
            "posters",
            "navy",
        }
        if hard_action_words.intersection(words):
            return False
        return len(rules_words.intersection(words)) <= 1

    def _fuzzy_name(
        self,
        name_hint: str,
        strict_words: bool = False,
        min_score: int = 84,
        prefer_en: bool = False,
    ) -> dict | None:
        normalized_hint = normalize_text(name_hint)
        if not normalized_hint:
            return None
        if process and fuzz:
            pt_card = None
            pt_score = 0
            pt_name = ""
            en_card = None
            en_score = 0
            en_name = ""
            pt_choices = self._pt_norm_choices()
            if pt_choices:
                pt_match = process.extractOne(normalized_hint, pt_choices.keys(), scorer=fuzz.WRatio)
                if pt_match and pt_match[1] >= min_score:
                    pt_name, pt_index = pt_choices[pt_match[0]]
                    pt_card = self.cards[pt_index]
                    pt_score = pt_match[1]
            choices = self._en_norm_choices()
            en_match = process.extractOne(normalized_hint, choices.keys(), scorer=fuzz.WRatio) if choices else None
            if en_match and en_match[1] >= min_score:
                candidate_name, candidate_index = choices[en_match[0]]
                if not strict_words or self._name_words_present(candidate_name, normalized_hint):
                    en_card = self.cards[candidate_index]
                    en_score = en_match[1]
                    en_name = candidate_name
            # A name that only explains part of what was read (e.g. the single-word
            # plane "Naya") must not win over a candidate that explains more of it
            # (e.g. "Medalhão de Naya"), even when English is normally preferred.
            if en_card and pt_card and en_card.get("oracle_id") != pt_card.get("oracle_id"):
                en_cov = self._hint_word_coverage(en_name, normalized_hint)
                pt_cov = self._hint_word_coverage(pt_name, normalized_hint)
                if pt_cov > en_cov:
                    return pt_card
                if en_cov > pt_cov:
                    return en_card
            if en_card and (not pt_card or en_score >= pt_score + 3 or prefer_en):
                return en_card
            if pt_card and (not en_card or pt_score > en_score + 3):
                return pt_card
            return en_card or pt_card

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

    @staticmethod
    def _hint_word_coverage(card_name: str, normalized_hint: str) -> int:
        """How many significant words of ``card_name`` appear in the read text.

        Used to decide, between two candidate names, which one better explains
        what the OCR actually read — so a short substring match cannot beat a
        fuller one.
        """
        name_words = [word for word in normalize_text(card_name).split() if len(word) >= 3]
        hint_words = [word for word in (normalized_hint or "").split() if len(word) >= 3]
        if not name_words or not hint_words:
            return 0
        return fuzzy_word_count(name_words, hint_words, threshold=85)

    def _fuzzy_portuguese_text(
        self,
        raw_text: str,
        *,
        namebar_mode: bool = False,
        max_lines: int = 20,
    ) -> tuple[dict | None, str]:
        if not raw_text or not process or not fuzz or not self.pt_name_choices:
            return None, ""
        if not namebar_mode:
            short_lines = []
            for line in raw_text.splitlines():
                cleaned = clean_ocr_line(line)
                words = normalize_text(cleaned).split()
                if not words or len(words) > 8:
                    continue
                short_lines.append(line)
                if len(short_lines) >= max_lines:
                    break
            if not short_lines:
                return None, ""
            raw_text = "\n".join(short_lines)
        raw_words = [word for word in normalize_text(raw_text).split() if len(word) >= 3]
        if namebar_mode:
            best_word_card = None
            best_word_name = ""
            best_word_score = 0
            for pt_name, index in self.pt_name_choices:
                pt_words = [word for word in normalize_text(pt_name).split() if len(word) >= 4]
                if not pt_words:
                    continue
                all_pt_words = [word for word in normalize_text(pt_name).split() if len(word) >= 3]
                first_word = all_pt_words[0] if all_pt_words else ""
                if len(all_pt_words) > 1 and first_word:
                    if not any(words_fuzzy_match(first_word, word, 88) for word in raw_words):
                        continue
                matches = fuzzy_word_count(pt_words, raw_words)
                required = min(3, len(pt_words)) if len(pt_words) >= 3 else min(2, len(pt_words))
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
            cleaned = clean_ocr_line(line)
            normalized = normalize_text(cleaned)
            words = [word for word in normalized.split() if len(word) >= 3]
            if not words or len(words) > 8:
                continue
            for size in range(min(5, len(words)), 1, -1):
                for start in range(0, len(words) - size + 1):
                    candidate = " ".join(words[start : start + size])
                    match = process.extractOne(candidate, choices.keys(), scorer=fuzz.WRatio)
                    if not match:
                        continue
                    pt_name = match[0]
                    pt_words = [word for word in normalize_text(pt_name).split() if len(word) >= 4]
                    all_pt_words = [word for word in normalize_text(pt_name).split() if len(word) >= 3]
                    first_word = all_pt_words[0] if all_pt_words else ""
                    if len(all_pt_words) > 1 and first_word:
                        if not any(words_fuzzy_match(first_word, word, 88) for word in words):
                            continue
                    if len(pt_words) == 1 and len(words) > 2:
                        continue
                    required = min(3, len(pt_words)) if len(pt_words) >= 3 else min(2, len(pt_words))
                    if pt_words and fuzzy_word_count(pt_words, words) < required:
                        continue
                    adjusted_score = match[1] + (len(pt_words) * 3)
                    if adjusted_score > best_score:
                        best_score = adjusted_score
                        best_name = pt_name
                        best_card = self.cards[choices[pt_name]]
        if best_card and best_score >= 90:
            return best_card, best_name
        return None, ""

    def _exact_english_title_in_text(self, raw_text: str) -> tuple[dict | None, str]:
        if not raw_text:
            return None, ""
        choices = self._en_norm_choices()
        lines_seen = 0
        ignored = COMMON_EN_OCR_WORDS | COMMON_PT_OCR_WORDS | {"again", "letter"}
        for line in raw_text.splitlines():
            cleaned = clean_ocr_line(line)
            words = [word for word in normalize_text(cleaned).split() if len(word) >= 3]
            if not words or len(words) > 3:
                continue
            lines_seen += 1
            if lines_seen > 80:
                break
            for size in range(min(5, len(words)), 0, -1):
                candidate = " ".join(words[:size])
                if candidate not in choices:
                    continue
                first_word = words[0]
                if len(first_word) < 5 or first_word in ignored:
                    continue
                original_name, card_index = choices[candidate]
                if len(normalize_text(original_name).split()) == size:
                    return self.cards[card_index], cleaned
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
        normalized_choices = self._en_norm_choices()
        best_card = None
        best_line = ""
        best_score = 0
        lines_seen = 0
        for line in raw_text.splitlines():
            cleaned = clean_ocr_line(line)
            normalized = normalize_text(cleaned)
            words = [word for word in normalized.split() if word not in ignored and len(word) > 1]
            if len(words) > 8:
                continue
            lines_seen += 1
            if lines_seen > 60:
                break
            if words:
                first_word = words[0]
                if len(words) <= 3 and len(first_word) >= 5 and first_word in normalized_choices:
                    original_name, card_index = normalized_choices[first_word]
                    if len(normalize_text(original_name).split()) == 1:
                        best_card = self.cards[card_index]
                        best_line = cleaned
                        best_score = max(best_score, 95)
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
        self.last_appended_row: int | None = None
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

        self.result_banner = Label(
            right,
            text="",
            font=("Segoe UI", 11, "bold"),
            anchor="w",
            padx=10,
            pady=8,
            wraplength=820,
            justify="left",
        )
        self.result_banner.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 10))

        for row_index, field in enumerate(FIELDS):
            ttk.Label(right, text=field).grid(row=row_index + 1, column=0, sticky="nw", padx=(0, 8), pady=3)
            entry = ttk.Entry(right, textvariable=self.field_vars[field])
            if field == "Descrição magia":
                entry = ttk.Entry(right, textvariable=self.field_vars[field])
            entry.grid(row=row_index + 1, column=1, sticky="ew", pady=3)

        ttk.Label(right, text="OCR bruto").grid(row=len(FIELDS) + 1, column=0, sticky="nw", padx=(0, 8), pady=3)
        self.raw_text = ttk.Label(right, textvariable=self.raw_ocr_var, wraplength=850, justify="left")
        self.raw_text.grid(row=len(FIELDS) + 1, column=1, sticky="ew", pady=3)

        status = ttk.Label(self.root, textvariable=self.status_var, anchor="w", padding=(10, 4))
        status.grid(row=1, column=0, columnspan=2, sticky="ew")

        self.duplicate_alert = Label(
            self.root,
            text="",
            fg="red",
            font=("Segoe UI", 10, "bold"),
            anchor="w",
            padx=10,
        )
        self.duplicate_alert.grid(row=2, column=0, columnspan=2, sticky="ew")

    def _log_ocr_status(self) -> None:
        if self.ocr.available():
            lang_note = "por+eng" if self.ocr.por_available else "eng (instale por no Tesseract para cartas PT)"
            self.log(f"OCR pronto: {self.ocr.tesseract_path} | {lang_note} | v{APP_VERSION}")
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
        self.last_appended_row = None
        self.clear_duplicate_alert()
        self.clear_result_banner()

    def show_result_banner(self, message: str, kind: str = "info") -> None:
        styles = {
            "success": ("#1b5e20", "#e8f5e9"),
            "duplicate": ("#b71c1c", "#ffebee"),
            "info": ("#1565c0", "#e3f2fd"),
        }
        fg, bg = styles.get(kind, styles["info"])
        self.result_banner.configure(text=message, fg=fg, bg=bg)
        self.root.update_idletasks()

    def clear_result_banner(self) -> None:
        self.result_banner.configure(text="", bg="SystemButtonFace", fg="SystemWindowText")

    def show_duplicate_alert(self, row_number: int) -> None:
        message = f"Carta já existe na linha {row_number} de {EXCEL_PATH.name} — não foi adicionada de novo."
        self.duplicate_alert.configure(text=message)
        self.show_result_banner(message, "duplicate")

    def clear_duplicate_alert(self) -> None:
        self.duplicate_alert.configure(text="")

    def check_duplicate_alert(self, name: str) -> None:
        if not name.strip():
            self.clear_duplicate_alert()
            return
        existing_row = find_name_row_in_excel(name)
        if existing_row is not None:
            self.show_duplicate_alert(existing_row)
        else:
            self.clear_duplicate_alert()

    def save_or_alert_duplicate(self, row: list[str], allow_update: bool = False) -> None:
        name = row[0]
        if not name.strip():
            return
        existing_row = find_name_row_in_excel(name)
        if existing_row is not None:
            if allow_update and self.last_appended_row == existing_row:
                update_row_in_excel(existing_row, row)
                self.clear_duplicate_alert()
                self.log(f"Atualizado na linha {existing_row} em {EXCEL_PATH}")
                return
            self.show_duplicate_alert(existing_row)
            return
        self.clear_duplicate_alert()
        append_row_to_excel(row)
        self.last_appended_row = find_name_row_in_excel(name)
        self.log(f"Salvo em {EXCEL_PATH}")
        if self.last_appended_row:
            self.show_result_banner(
                f"Salvo na linha {self.last_appended_row} de {EXCEL_PATH.name}.",
                "success",
            )

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
        self.last_appended_row = None
        self.clear_results()
        self.raw_ocr_var.set(f"Extraindo dados... v{APP_VERSION}")
        self.set_busy(True, "Extraindo dados da carta...")

        def worker() -> None:
            ocr_result = None
            try:
                self.log("Carregando base local...")
                self.db.load(log=self.log)
                self.log("Rodando OCR local...")
                ocr_result = self.ocr.extract(image_for_job)
                self.log("Procurando carta na base...")
                card, reason = self.db.find(ocr_result)
                if job_id != self.extraction_id:
                    return
                self.root.after(0, lambda: self._finish_extract(job_id, ocr_result, card, reason))
            except Exception as exc:
                if job_id != self.extraction_id:
                    return
                message = str(exc)
                captured_ocr = ocr_result
                self.root.after(0, lambda: self._fail_extract(job_id, message, captured_ocr))

        threading.Thread(target=worker, daemon=True).start()

    def _fail_extract(self, job_id: int, message: str, ocr_result: OcrResult | None = None) -> None:
        if job_id != self.extraction_id:
            return
        self.clear_results()
        self.raw_ocr_var.set(f"Erro: {message}")
        self.log(f"Erro: {message}")
        if ocr_result is not None:
            self.write_debug({}, "", ocr_result)
        messagebox.showerror("Erro", message)
        self.set_busy(False)

    def _finish_extract(self, job_id: int, ocr_result: OcrResult, card: dict, reason: str) -> None:
        if job_id != self.extraction_id:
            return
        try:
            self.write_debug(card, reason, ocr_result)
            row = self.db.to_row(card, self.foil_var.get())
            self.fill_fields(row, ocr_result, reason, job_id)
            existing_row = find_name_row_in_excel(row[0])
            if existing_row is None:
                self.save_or_alert_duplicate(self.row_values())
                saved_row = find_name_row_in_excel(row[0])
                if saved_row:
                    message = f"Salvo na linha {saved_row} de {EXCEL_PATH.name}."
                    self.show_result_banner(message, "success")
                    self.log(message)
            else:
                self.show_duplicate_alert(existing_row)
                self.log(
                    f"Preenchido: {row[0]} | {reason} | "
                    f"Já existe na linha {existing_row} de {EXCEL_PATH.name} — não foi adicionada de novo"
                )
                messagebox.showwarning(
                    "Carta já existe",
                    f"\"{row[0]}\" já está na linha {existing_row} de {EXCEL_PATH.name}.\n\n"
                    "Os campos foram preenchidos, mas a carta não foi adicionada de novo ao Excel.",
                )
        except Exception as exc:
            self._fail_extract(job_id, str(exc), ocr_result)
            return
        self.set_busy(False)

    def write_debug(self, card: dict, reason: str, ocr_result: OcrResult) -> None:
        lines = [
            f"version: {APP_VERSION}",
            f"timestamp: {time.strftime('%Y-%m-%d %H:%M:%S')}",
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
        text = "\n".join(lines)
        DEBUG_PATH.write_text(text, encoding="utf-8")
        # Keep a rolling history so failed extractions can be reviewed later
        # (the single-file debug above is overwritten on every extraction).
        try:
            with DEBUG_LOG_PATH.open("a", encoding="utf-8") as handle:
                handle.write(f"\n{'=' * 70}\n{text}\n")
        except Exception:
            pass

    def fill_fields(self, row: list[str], ocr_result: OcrResult, reason: str, job_id: int) -> None:
        if job_id != self.extraction_id:
            return
        self.clear_results()
        for field, value in zip(FIELDS, row):
            self.field_vars[field].set(value)
        raw = ocr_result.raw_text.strip()
        self.raw_ocr_var.set(raw[:1200])
        self.show_result_banner(f"Carta identificada: {row[0]}", "info")
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
        self.save_or_alert_duplicate(row, allow_update=True)

    def download_database(self) -> None:
        self.run_background(lambda: self.db.download(self.log), "Base atualizada.")


def main() -> None:
    root = Tk()
    app = MagicExtractorApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
