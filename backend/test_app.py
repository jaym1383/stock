import tempfile
import unittest
from io import BytesIO
from pathlib import Path
from unittest.mock import patch

from openpyxl import Workbook

import app as stock_app


class SharedStockTests(unittest.TestCase):
    def setUp(self):
        self.directory = tempfile.TemporaryDirectory()
        self.original_path = stock_app.STOCK_PATH
        stock_app.STOCK_PATH = Path(self.directory.name) / "stock.xlsx"
        self.first_client = stock_app.app.test_client()
        self.second_client = stock_app.app.test_client()

    def tearDown(self):
        stock_app.STOCK_PATH = self.original_path
        self.directory.cleanup()

    @staticmethod
    def workbook_bytes(quantity):
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "Phones"
        sheet.append(["Model", "Storage", "Color", "Qty"])
        sheet.append(["Test phone", "128GB", "Black", quantity])
        output = BytesIO()
        workbook.save(output)
        return output.getvalue()

    def upload(self, content):
        return self.first_client.post(
            "/api/upload-stock",
            data={"file": (BytesIO(content), "stock.xlsx")},
            content_type="multipart/form-data",
            headers={"Origin": "https://jaym1383.github.io"},
        )

    def shared_stock_bytes(self):
        response = self.second_client.get("/api/stock")
        try:
            return response.data
        finally:
            response.close()

    def test_other_client_sees_new_upload_and_version(self):
        self.assertIsNone(self.second_client.get("/api/stock-version").json["version"])
        self.assertEqual(self.second_client.get("/api/stock").status_code, 404)

        first = self.upload(self.workbook_bytes(2))
        self.assertEqual(first.status_code, 200)
        first_version = first.json["version"]
        self.assertEqual(first.headers["Access-Control-Allow-Origin"], "https://jaym1383.github.io")
        self.assertIn("X-Stock-Version", first.headers["Access-Control-Expose-Headers"])
        self.assertEqual(self.second_client.get("/api/stock-version").json["version"], first_version)
        self.assertEqual(self.shared_stock_bytes(), self.workbook_bytes(2))

        second = self.upload(self.workbook_bytes(5))
        self.assertEqual(second.status_code, 200)
        self.assertNotEqual(second.json["version"], first_version)
        self.assertEqual(self.shared_stock_bytes(), self.workbook_bytes(5))

    def test_invalid_upload_keeps_previous_stock(self):
        valid = self.workbook_bytes(3)
        self.assertEqual(self.upload(valid).status_code, 200)
        rejected = self.upload(b"not an xlsx file")
        self.assertEqual(rejected.status_code, 400)
        self.assertEqual(self.shared_stock_bytes(), valid)

    def test_pdf_upload_publishes_shared_stock(self):
        workbook = self.workbook_bytes(7)
        rows = [["Phones", "Test phone", "128GB", "Black", 7, "CODE"]]

        def write_mock(_rows, output_path):
            Path(output_path).write_bytes(workbook)

        with patch("pdf_to_stock.extract_records", return_value=[]), \
             patch("pdf_to_stock.build_rows", return_value=rows), \
             patch("pdf_to_stock.write_workbook", side_effect=write_mock):
            response = self.first_client.post(
                "/api/convert",
                data={"file": (BytesIO(b"%PDF-1.4"), "stock.pdf")},
                content_type="multipart/form-data",
            )

        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.headers["X-Stock-Rows"], "1")
        self.assertEqual(response.headers["X-Stock-Quantity"], "7")
        self.assertEqual(response.headers["X-Stock-Version"], self.second_client.get("/api/stock-version").json["version"])
        self.assertEqual(self.shared_stock_bytes(), workbook)


if __name__ == "__main__":
    unittest.main()
