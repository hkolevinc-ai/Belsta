import importlib.util
import sys
import unittest
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[1]
SPEC = importlib.util.spec_from_file_location("belsta_scraper", PROJECT_ROOT / "scraper.py")
assert SPEC and SPEC.loader
SCRAPER = importlib.util.module_from_spec(SPEC)
sys.modules[SPEC.name] = SCRAPER
SPEC.loader.exec_module(SCRAPER)


class RetailExpansionTests(unittest.TestCase):
    def test_size_range_and_multiplicity(self):
        self.assertEqual(SCRAPER.expand_size_range("23-27 2*26"), ["23", "24", "25", "26", "27"])
        self.assertEqual(SCRAPER.count_pairs("23-27 2*26"), 6)
        self.assertEqual(SCRAPER.expand_size_range("36"), ["36"])

    def test_wholesale_variants_become_unique_retail_rows(self):
        product = {
            "id": 100,
            "handle": "test-pack",
            "title": "Опаковка: Детски пантофи BELSTA на едро",
            "description": "",
            "options": [{"name": "Размер"}, {"name": "Цвят"}],
            "images": ["https://example.com/blue.jpg"],
            "variants": [
                {
                    "id": 101,
                    "available": True,
                    "option1": "23-27 2*26",
                    "option2": "Син",
                    "price": 6000,
                    "compare_at_price": 7200,
                    "weight": 3000,
                    "featured_image": {"src": "https://example.com/blue.jpg"},
                },
                {
                    "id": 102,
                    "available": True,
                    "option1": "23-27",
                    "option2": "Син",
                    "price": 5500,
                    "compare_at_price": None,
                    "weight": 2500,
                    "featured_image": {"src": "https://example.com/blue.jpg"},
                },
            ],
        }
        selection = SCRAPER.InputSelection(
            row_number=2,
            original_url="https://www.belsta.bg/products/test-pack?variant=101",
            handle="test-pack",
            variant_id="101",
        )

        items, warnings = SCRAPER.select_variants([selection], {"test-pack": product})

        self.assertEqual(warnings, [])
        self.assertEqual([item.size for item in items], ["23", "24", "25", "26", "27"])
        self.assertTrue(all(item.variant["id"] == 101 for item in items))
        self.assertTrue(all(item.source_pack_pairs == 6 for item in items))

        rows = SCRAPER.build_workbook_rows(items)
        self.assertEqual([row["KI"] for row in rows], ["23", "24", "25", "26", "27"])
        self.assertTrue(all(row["L"] == "Детски пантофи BELSTA" for row in rows))
        self.assertTrue(all(row["LZ"] == 10.0 for row in rows))
        self.assertTrue(all(row["MB"] == 12.0 for row in rows))
        self.assertTrue(all(row["MD"] == 500 for row in rows))
        self.assertTrue(all(row["MH"] == "Single set" for row in rows))
        self.assertTrue(all(row["ML"] == 1 for row in rows))
        self.assertEqual(len({row["N"] for row in rows}), 5)


if __name__ == "__main__":
    unittest.main()
