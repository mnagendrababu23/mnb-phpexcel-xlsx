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
            'positioning' => 'MNB PHPExcel provides native application-ready Excel import/export workflows while keeping heavyweight calculation and legacy-format engines optional.',
            'capabilities' => [
                'deep_cell_level_editing' => [
                    'status' => 'supported_for_read_and_generated_workbooks',
                    'current' => ['single-cell reads', 'multi-cell reads', 'range reads', 'typed cell snapshots', 'complete style metadata', 'typed rich text'],
                    'limits' => ['existing complex workbooks are not exposed as a mutable random-access object graph'],
                ],
                'many_spreadsheet_formats' => [
                    'status' => 'supported_modularly',
                    'current' => ['native XLSX', 'CSV/TSV', 'JSON/NDJSON', 'XML', 'native ODS reading', 'optional PhpSpreadsheet-backed XLS reading'],
                    'limits' => ['legacy XLS writing and ODS writing are not part of the current package family'],
                ],
                'advanced_style_manipulation' => [
                    'status' => 'supported',
                    'current' => ['fonts', 'fills', 'borders', 'alignment', 'protection', 'number formats', 'cell/range/row/column styles', 'native conditional formatting', 'validated templates'],
                    'limits' => ['large streaming mode intentionally exposes a smaller style surface for bounded memory usage'],
                ],
                'formula_calculation_engine' => [
                    'status' => 'supported_by_optional_adapter',
                    'current' => ['formula expression reads', 'cached-result reads', 'formula writing', 'recalculate-on-open flags', 'optional PhpSpreadsheet recalculation adapter'],
                    'limits' => ['the lightweight native core does not ship its own calculation engine'],
                ],
                'charts' => [
                    'status' => 'supported',
                    'current' => ['column', 'bar', 'line', 'area', 'pie', 'doughnut', 'scatter', 'drawing relationships and package parts'],
                    'limits' => ['advanced chart families and arbitrary existing-chart mutation remain template/adapter workflows'],
                ],
                'pivot_tables' => [
                    'status' => 'supported_template_workflow',
                    'current' => ['pivot table preservation', 'pivot cache preservation', 'source range rebinding', 'refresh-on-open', 'related package-part preservation'],
                    'limits' => ['from-scratch pivot layout design is intentionally not advertised as supported'],
                ],
                'complex_workbook_preservation' => [
                    'status' => 'supported_for_trusted_templates',
                    'current' => ['comments', 'hyperlinks', 'drawings', 'images', 'charts', 'pivot parts', 'tables', 'external links', 'integrity validation'],
                    'limits' => ['macros are never executed; deep arbitrary package-graph editing remains an adapter concern'],
                ],
            ],
            'adapter_recommendations' => [
                'php_spreadsheet_adapter' => 'Optional dependency for true formula recalculation and legacy XLS compatibility.',
                'framework_adapters' => ['Laravel', 'CodeIgniter', 'Slim/Symfony integrations should remain separate packages'],
            ],
        ];
    }
}
