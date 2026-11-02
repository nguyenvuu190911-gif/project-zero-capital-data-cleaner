# Data Cleaner — CSV & simple Excel (.xlsx)

Local utility for cleaning tabular files. Python 3.9+; no paid API or third-party package.

```powershell
python data_cleaner.py sample_before.csv cleaned.csv --report report.json
python data_cleaner.py input.xlsx cleaned.xlsx --report report.json
python -m unittest discover -s . -p test_data_cleaner.py -v
```

Trims surrounding whitespace, removes fully blank and exact-duplicate rows, counts rows with extra columns as malformed, and preserves missing cells as blanks. It writes a separate report and never overwrites the source. Both sample CSV files contain invented data.

XLSX support is limited to one worksheet of plain values. Output is a new value-only workbook; styles, formulas, charts, comments, date formats, and extra sheets are not preserved. Formula cells and multiple sheets are rejected. Back up originals and inspect the report. Avoid huge untrusted workbooks and use client data only with permission.

MIT licensed. No telemetry or network calls.

## Dịch vụ / Service

**VI:** Làm sạch bảng CSV/Excel nhỏ: bỏ dòng trống và trùng hoàn toàn, xử lý khoảng trắng, đối chiếu số dòng và bàn giao báo cáo. Gửi cấu trúc cột và mẫu đã ẩn dữ liệu riêng tư để nhận báo giá cố định. Quy tắc ngày, số, trùng gần đúng hoặc nhiều sheet được báo giá riêng.

**EN:** Small CSV/Excel cleanup with a row-count audit: trim whitespace, remove empty and exact-duplicate rows, and deliver a separate change report. Share a redacted sample and column structure for a fixed quote. Date/number normalization, fuzzy matching, and multi-sheet work require a separate scope.
