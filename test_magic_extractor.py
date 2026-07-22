import unittest

from magic_extractor import OcrResult, ScryfallDatabase


class EnglishTitleRegressionTests(unittest.TestCase):
    def setUp(self) -> None:
        self.database = ScryfallDatabase()
        self.database.cards = [
            {
                "name": "Barrage of Expendables",
                "oracle_id": "barrage",
                "lang": "en",
            },
            {
                "name": "Goblin General",
                "oracle_id": "goblin-general",
                "lang": "en",
            },
        ]
        self.database.name_choices = [
            ("Barrage of Expendables", 0),
            ("Goblin General", 1),
        ]

    def test_stray_lowercase_letter_before_title_is_ignored(self) -> None:
        namebar_text = """fearrase of Expendables —

; Barrage of Exper

a Barrage of Expendables ‘
"""

        card, matched_line = self.database._exact_english_title_in_text(namebar_text)

        self.assertIsNotNone(card)
        self.assertEqual("Barrage of Expendables", card["name"])
        self.assertEqual("a Barrage of Expendables ‘", matched_line)


class PortugueseTitleRegressionTests(unittest.TestCase):
    def test_clean_title_beats_higher_ranked_noise_line(self) -> None:
        database = ScryfallDatabase()
        card = {
            "name": "Stoneshock Giant",
            "oracle_id": "stoneshock-giant",
            "lang": "en",
        }
        database.mtgjson_pt_by_name["stoneshock giant"] = {
            "printed_name": "Gigante Petrochoque",
        }
        ocr = OcrResult(
            namebar_text="""{(Gigance Perrochogue = 3

: Ties = erro, A mene . > er a a

m Gigante Petrochoque 38
""",
        )

        matched = database._strong_portuguese_namebar_match(
            card, "Gigante Petrochoque", ocr
        )

        self.assertTrue(matched)


if __name__ == "__main__":
    unittest.main()
