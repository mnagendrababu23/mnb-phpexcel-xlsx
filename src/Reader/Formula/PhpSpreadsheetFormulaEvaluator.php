<?php

declare(strict_types=1);

namespace Mnb\PHPExcel\Reader\Formula;

use Mnb\PHPExcel\Support\Coordinate;
use Mnb\PHPExcel\Support\MnbExcelException;

final class PhpSpreadsheetFormulaEvaluator implements FormulaEvaluatorInterface
{
    public function available(): bool
    {
        return class_exists('PhpOffice\\PhpSpreadsheet\\IOFactory');
    }

    public function calculateCell(string $path, int|string $sheet, string $cell): mixed
    {
        $spreadsheet = $this->load($path);
        try {
            $worksheet = $this->worksheet($spreadsheet, $sheet);
            return $worksheet->getCell(strtoupper($cell))->getCalculatedValue();
        } finally {
            if (method_exists($spreadsheet, 'disconnectWorksheets')) {
                $spreadsheet->disconnectWorksheets();
            }
        }
    }

    /** @return array<string,mixed> */
    public function calculateRange(string $path, int|string $sheet, string $range): array
    {
        [$start, $end] = $this->normalizeRange($range);
        [$startColumn, $startRow] = Coordinate::splitCellRef($start);
        [$endColumn, $endRow] = Coordinate::splitCellRef($end);

        $spreadsheet = $this->load($path);
        try {
            $worksheet = $this->worksheet($spreadsheet, $sheet);
            $values = [];
            for ($row = min($startRow, $endRow); $row <= max($startRow, $endRow); $row++) {
                for ($column = min($startColumn, $endColumn); $column <= max($startColumn, $endColumn); $column++) {
                    $cell = Coordinate::columnIndexToName($column) . $row;
                    $values[$cell] = $worksheet->getCell($cell)->getCalculatedValue();
                }
            }
            return $values;
        } finally {
            if (method_exists($spreadsheet, 'disconnectWorksheets')) {
                $spreadsheet->disconnectWorksheets();
            }
        }
    }

    private function load(string $path): object
    {
        if (!$this->available()) {
            throw new MnbExcelException(
                'True formula recalculation requires the optional phpoffice/phpspreadsheet package. '
                . 'Install it with: composer require phpoffice/phpspreadsheet'
            );
        }
        if (!is_file($path)) {
            throw new MnbExcelException('Spreadsheet file not found: ' . $path);
        }

        /** @var class-string $ioFactory */
        $ioFactory = 'PhpOffice\\PhpSpreadsheet\\IOFactory';
        return $ioFactory::load($path);
    }

    private function worksheet(object $spreadsheet, int|string $sheet): object
    {
        if (is_int($sheet) || ctype_digit((string) $sheet)) {
            $index = max(0, (int) $sheet - 1);
            return $spreadsheet->getSheet($index);
        }

        $worksheet = $spreadsheet->getSheetByName((string) $sheet);
        if ($worksheet === null) {
            throw new MnbExcelException('Worksheet not found: ' . $sheet);
        }
        return $worksheet;
    }

    /** @return array{string,string} */
    private function normalizeRange(string $range): array
    {
        $range = strtoupper(trim($range));
        if (preg_match('/^([A-Z]+\d+)(?::([A-Z]+\d+))?$/', $range, $match) !== 1) {
            throw new MnbExcelException('Invalid cell range: ' . $range);
        }
        return [$match[1], $match[2] ?? $match[1]];
    }
}
