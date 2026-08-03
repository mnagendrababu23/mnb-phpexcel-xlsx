# Visual Snapshot and Style Round-Trip

The visual snapshot API captures spreadsheet values and the presentation information needed to recreate a workbook without application-specific style conversion code.

## Public API

```php
use Mnb\PHPExcel\MnbExcel;
use Mnb\PHPExcel\Snapshot\VisualSnapshot;

$snapshot = MnbExcel::visualSnapshot('source.xlsx');
$json = VisualSnapshot::toJson($snapshot);
file_put_contents('source.visual.json', $json);

MnbExcel::createFromVisualSnapshot(
    file_get_contents('source.visual.json'),
    'recreated.xlsx'
);
```

Format packages expose the same methods:

```php
Xlsx::visualSnapshot($path);
Xlsx::createFromVisualSnapshot($snapshotOrJson, $destination);

Xls::visualSnapshot($path);
Xls::createFromVisualSnapshot($snapshotOrJson, $destination);

Csv::visualSnapshot($path);
Csv::createFromVisualSnapshot($snapshotOrJson, $destination);
```

Reading sessions also expose the operation:

```php
$snapshot = Xlsx::read($path)->visualSnapshot();
```

Password-protected XLSX files use the same options on read and write:

```php
$snapshot = Xlsx::visualSnapshot('protected.xlsx', [
    'password' => $password,
]);

Xlsx::createFromVisualSnapshot($snapshot, 'protected-copy.xlsx', [
    'password' => $password,
    'encryption_mode' => 'standard', // or agile
]);
```

## Schema

Every snapshot has the following stable envelope:

```php
[
    'schema' => 'mnb-phpexcel.visual-snapshot',
    'schema_version' => '1.0',
    'format' => 'xlsx',
    'source' => [],
    'workbook' => [],
    'styles' => [],
    'sheets' => [],
    'capabilities' => [],
    'warnings' => [],
]
```

Cells are stored sparsely by coordinate instead of as a large rectangular matrix:

```php
'cells' => [
    'A1' => [
        'type' => 'text',
        'value' => 'Order ID',
        'style_id' => 's1',
    ],
    'B2' => [
        'type' => 'date',
        'value' => '2026-07-01T00:00:00+00:00',
        'format' => 'yyyy-mm-dd',
        'style_id' => 's2',
    ],
]
```

Sparse coordinates prevent accidental column loss and avoid exporting thousands of empty style entries.

## Captured workbook features

### XLSX

- Text, numbers, booleans, errors, blanks, formulas, cached formula values and typed dates
- Shared and inline rich text
- Fonts, fills, borders, alignments, number formats and cell protection
- Column widths and row heights
- Merged cells
- Frozen rows and columns
- Auto-filter range
- Conditional color scales, data bars, icon sets and supported cell rules
- Data validation rules
- Comments and hyperlinks
- Image inventory
- Active sheet and hidden/very-hidden sheet states
- Workbook date system and document metadata

### XLS/BIFF8

- Cell values, formulas and true numeric dates
- BIFF8 fonts, palette colors, pattern fills, borders, alignment, number formats and protection
- Column widths and row heights
- Merged cells and frozen panes
- Active sheet and hidden/very-hidden sheet states
- Workbook date system

Unsupported BIFF8 features are reported through the snapshot capability fields instead of being silently presented as restored.

### CSV/TSV

CSV has no embedded styles or workbook layout. The snapshot preserves:

- Values and column positions
- Encoding and BOM
- Delimiter, enclosure, escape character and line ending

The snapshot reports styles and layout as `not_applicable`.

## Canonical style shape

The reader and writers use one canonical style structure:

```php
[
    'font' => [
        'name' => 'Calibri',
        'size' => 11,
        'bold' => true,
        'italic' => false,
        'color' => ['rgb' => 'FFFFFFFF'],
    ],
    'fill' => [
        'type' => 'pattern',
        'pattern' => 'solid',
        'foreground' => ['rgb' => 'FF1F4E78'],
    ],
    'border' => [
        'left' => ['style' => 'thin', 'color' => ['rgb' => 'FF808080']],
        'right' => ['style' => 'thin', 'color' => ['rgb' => 'FF808080']],
        'top' => ['style' => 'thin', 'color' => ['rgb' => 'FF808080']],
        'bottom' => ['style' => 'thin', 'color' => ['rgb' => 'FF808080']],
    ],
    'alignment' => [
        'horizontal' => 'center',
        'vertical' => 'center',
        'wrap_text' => false,
    ],
    'format' => 'yyyy-mm-dd',
]
```

Writer input aliases remain accepted. For example, this documented shape is normalized correctly:

```php
[
    'borders' => [
        'all' => [
            'style' => 'thin',
            'color' => '#808080',
        ],
    ],
]
```

## Real Excel dates

`dateStyleColumns()` and `datetimeStyleColumns()` in the Mono report builder now create typed date cells when input values are parseable dates. The writer stores an Excel serial number and applies the requested number format.

```php
MnbExcel::report($rows)
    ->dateStyleColumns(['Order Date'], 'yyyy-mm-dd')
    ->save('report.xlsx');
```

Use `MnbExcel::date()` or `CellValue::date()` when explicit per-cell typing is preferable.

## Capability rules

A snapshot never implies that every destination format supports every feature. Inspect:

```php
$snapshot['capabilities'];
$snapshot['sheets'][0]['capabilities'];
$snapshot['warnings'];
```

When converting XLSX to XLS, OOXML-only conditional formatting or validation details may not be writable. When converting to CSV, all presentation features are intentionally omitted.

## Safety and limits

- Formula cells remain explicit typed cells.
- Formula-like plain text remains text when the workbook enables formula escaping.
- Snapshot input is validated before writing.
- `max_cells` defaults to 1,000,000 when hydrating a snapshot.
- Unsupported image inventory is not treated as embedded image bytes.
- No macro, external link, ActiveX control or embedded object is executed.
