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

    def test_link_without_variant_uses_only_default_color(self):
        product = {
            "id": 200,
            "handle": "default-color",
            "title": "Дамски пантофи",
            "description": "",
            "options": [{"name": "Размер"}, {"name": "Цвят"}],
            "images": ["https://example.com/pink.jpg", "https://example.com/blue.jpg"],
            "variants": [
                {
                    "id": 201,
                    "available": True,
                    "option1": "36",
                    "option2": "Розов",
                    "price": 1000,
                    "weight": 500,
                    "featured_image": {"src": "https://example.com/pink.jpg"},
                },
                {
                    "id": 202,
                    "available": True,
                    "option1": "37",
                    "option2": "Розов",
                    "price": 1000,
                    "weight": 500,
                    "featured_image": {"src": "https://example.com/pink.jpg"},
                },
                {
                    "id": 203,
                    "available": True,
                    "option1": "36",
                    "option2": "Син",
                    "price": 1000,
                    "weight": 500,
                    "featured_image": {"src": "https://example.com/blue.jpg"},
                },
            ],
        }
        selection = SCRAPER.InputSelection(
            row_number=2,
            original_url="https://www.belsta.bg/products/default-color",
            handle="default-color",
            variant_id=None,
        )

        items, warnings = SCRAPER.select_variants([selection], {"default-color": product})

        self.assertEqual(warnings, [])
        self.assertEqual([item.size for item in items], ["36", "37"])
        self.assertEqual({item.source_color for item in items}, {"Розов"})

    def test_color_collisions_are_disambiguated(self):
        women = SCRAPER.assign_temu_colors(
            ["Розов", "Флорален"], SCRAPER.CATEGORY_WOMEN_SLIPPERS
        )
        blues = SCRAPER.assign_temu_colors(
            ["Blue", "Син"], SCRAPER.CATEGORY_GIRLS_SLIPPERS
        )

        self.assertEqual(women, {"Розов": "Pink", "Флорален": "Deep Pink"})
        self.assertEqual(blues, {"Blue": "Blue", "Син": "Royal Blue"})

    def test_wholesale_bundle_images_are_blocked(self):
        blocked = (
            "https://cdn.shopify.com/files/"
            "84b23fa3b996d126681d2788ff34ba9c_example.jpg?v=1"
        )
        blocked_exact = (
            "https://cdn.shopify.com/files/"
            "1_51684791-009f-4c11-9d76-741ed70bcc42.png?v=1"
        )
        normal = "https://cdn.shopify.com/files/retail_pair.jpg?v=1"

        self.assertTrue(SCRAPER.is_wholesale_image_url(blocked))
        self.assertTrue(SCRAPER.is_wholesale_image_url(blocked_exact))
        self.assertFalse(SCRAPER.is_wholesale_image_url(normal))


if __name__ == "__main__":
    unittest.main()
