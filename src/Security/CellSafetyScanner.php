<?php

declare(strict_types=1);

namespace Mnb\PHPExcel\Security;

use Mnb\PHPExcel\Core\CellValue;
use Mnb\PHPExcel\Support\ValueSanitizer;

final class CellSafetyScanner
{
    /**
     * Scan array rows before export/import for common cell risks.
     *
     * @param list<array<int|string,mixed>> $rows
     * @param array<string,mixed> $options
     * @return array{status:string,total_issues:int,issues:list<array{row:int,column:string,type:string,message:string,value?:mixed}>}
     */
    public function scan(array $rows, array $options = []): array
    {
        $issues = [];
        $maxLength = (int) ($options['max_text_length'] ?? 32767);

        foreach ($rows as $rowIndex => $row) {
            if (!is_array($row)) {
                continue;
            }

            foreach ($row as $column => $value) {
                $columnName = (string) $column;
                if ($value instanceof CellValue) {
                    if ($value->type() === CellValue::TYPE_FORMULA) {
                        $risk = FormulaGuard::risk((string) $value->value(), $options);
                        if ($risk['risk'] !== 'none') {
                            $issues[] = [
                                'row' => $rowIndex + 1,
                                'column' => $columnName,
                                'type' => 'unsafe_formula',
                                'message' => implode('; ', $risk['reasons']),
                                'value' => '=' . (string) $value->value(),
                            ];
                        }
                    }
                    $value = $value->displayValue();
                }

                if (!is_string($value)) {
                    continue;
                }

                if (ValueSanitizer::isFormulaLikeText($value)) {
                    $issues[] = [
                        'row' => $rowIndex + 1,
                        'column' => $columnName,
                        'type' => 'formula_like_text',
                        'message' => 'Text starts with a formula trigger character.',
                        'value' => substr($value, 0, 120),
                    ];
                }
                if (ValueSanitizer::containsInvalidXmlCharacters($value)) {
                    $issues[] = [
                        'row' => $rowIndex + 1,
                        'column' => $columnName,
                        'type' => 'invalid_xml_characters',
                        'message' => 'Text contains characters that cannot be written safely into XLSX XML.',
                    ];
                }
                if ($maxLength > 0 && ValueSanitizer::textLength($value) > $maxLength) {
                    $issues[] = [
                        'row' => $rowIndex + 1,
                        'column' => $columnName,
                        'type' => 'text_too_long',
                        'message' => 'Text exceeds Excel cell text limit.',
                    ];
                }
                if (ValueSanitizer::isLargeIntegerString($value)) {
                    $issues[] = [
                        'row' => $rowIndex + 1,
                        'column' => $columnName,
                        'type' => 'large_number_text',
                        'message' => 'Long numeric text should be exported as text to avoid Excel precision loss.',
                        'value' => substr($value, 0, 80),
                    ];
                }
            }
        }

        return [
            'status' => $issues === [] ? 'ok' : 'warning',
            'total_issues' => count($issues),
            'issues' => $issues,
        ];
    }
}
