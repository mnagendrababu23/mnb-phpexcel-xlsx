<?php

declare(strict_types=1);

namespace Mnb\PHPExcel\Support;

final class AdvancedWorkbookCapabilities
{
    /** @return array<string,mixed> */
    public static function matrix(): array
    {
        return [
            'status' => 'ready',
            'positioning' => 'MNB PHPExcel focuses on application-ready import/export workflows. Advanced spreadsheet manipulation is tracked explicitly so developers know what is supported now and what requires a specialist library or future adapter.',
            'capabilities' => [
                'deep_cell_level_editing' => [
                    'status' => 'partial',
                    'current' => ['write values', 'write formulas with cached values', 'styles by cell/range/row/column in normal mode', 'comments/hyperlinks write', 'merged cells'],
                    'planned' => ['random-access existing XLSX cell editing', 'preserve/rewrite arbitrary workbook part graph', 'cell comments edit-in-place'],
                    'recommendation' => 'Use current normal writer for generated reports; use future adapter/PhpSpreadsheet for deep editing of existing complex workbooks.',
                ],
                'many_spreadsheet_formats' => [
                    'status' => 'partial',
                    'current' => ['XLSX', 'CSV', 'JSON', 'XML', 'CSV ZIP fallback'],
                    'planned' => ['ODS adapter', 'HTML table import/export adapter', 'NDJSON/Parquet style data export adapters'],
                    'recommendation' => 'Keep core lightweight; add extra formats through reader/writer plugin interfaces.',
                ],
                'advanced_style_manipulation' => [
                    'status' => 'partial',
                    'current' => ['headers', 'named styles', 'number formats', 'range/cell/row/column styles', 'basic large-writer formats'],
                    'planned' => ['conditional formatting writer', 'theme palettes', 'style copy from template', 'richer border/fill/font combinations'],
                    'recommendation' => 'Normal mode can be rich; large mode should stay style-limited for memory safety.',
                ],
                'formula_calculation_engine' => [
                    'status' => 'not_in_core',
                    'current' => ['formula writing', 'formula safety', 'formula cached value read/write'],
                    'planned' => ['optional calculation adapter interface', 'formula dependency report', 'recalculate-on-open workbook flag'],
                    'recommendation' => 'Do not build a full calculation engine in v1 core; use cached values or external specialist adapter.',
                ],
                'complex_workbook_manipulation' => [
                    'status' => 'partial_preservation',
                    'current' => ['metadata extraction', 'comments/hyperlinks', 'template advanced object preservation path', 'integrity validation'],
                    'planned' => ['complex workbook diff/report', 'template part dependency validator', 'adapter-based advanced editing'],
                    'recommendation' => 'Preserve advanced objects safely; avoid pretending to edit pivots/charts/macros directly in core.',
                ],
            ],
            'adapter_recommendations' => [
                'php_spreadsheet_adapter' => 'For deep workbook manipulation and formula/calculation-heavy workflows.',
                'openspout_adapter' => 'For alternate proven streaming read/write workflows.',
                'framework_adapters' => ['Laravel', 'CodeIgniter', 'Slim/Symfony examples'],
            ],
        ];
    }
}
