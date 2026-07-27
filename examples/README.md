# Lightweight XLSX information examples

These examples inspect an XLSX package without loading workbook rows or decoding
cell values.

```bash
php examples/file_info.php storage/orders.xlsx
php examples/sheet_info.php storage/orders.xlsx
php examples/row_count.php storage/orders.xlsx Orders
```

Accurate row counts stream worksheet XML when `ext-xmlreader` is available. The
pure-PHP fallback keeps the same API, but may buffer an individual worksheet XML
part internally.
