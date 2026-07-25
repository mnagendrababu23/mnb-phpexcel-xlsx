<?php

declare(strict_types=1);

namespace Mnb\PHPExcel\Reader;

use Mnb\PHPExcel\Support\Coordinate;
use Mnb\PHPExcel\Support\ErrorCode;
use Mnb\PHPExcel\Support\MnbExcelException;
use ZipArchive;

/** @internal XLSX merged-range expansion helper. */
final class MergedCellMap
{
    /** @param list<array{start_row:int,end_row:int,start_column:int,end_column:int,ref:string}> $ranges */
    private function __construct(private readonly array $ranges)
    {
    }

    /** @param array<string,mixed> $options */
    public static function fromXlsx(string $path, string $sheetEntry, array $options = []): self
    {
        $zip = new ZipArchive();
        if ($zip->open($path) !== true) {
            throw MnbExcelException::withCode('Unable to open XLSX while reading merged cells.', ErrorCode::XLSX_INVALID);
        }
        try {
            $xml = $zip->getFromName($sheetEntry);
        } finally {
            $zip->close();
        }
        if (!is_string($xml) || $xml === '') {
            return new self([]);
        }

        $ranges = [];
        $maxRanges = max(0, (int) ($options['max_merged_ranges'] ?? 100000));
        preg_match_all('/<mergeCell\b[^>]*\bref\s*=\s*(?:"([^"]+)"|\'([^\']+)\')[^>]*\/?\s*>/i', $xml, $matches, PREG_SET_ORDER);
        foreach ($matches as $match) {
            $ref = (string) (($match[1] ?? '') !== '' ? $match[1] : ($match[2] ?? ''));
            $parts = explode(':', str_replace('$', '', $ref), 2);
            if (count($parts) !== 2) {
                continue;
            }
            try {
                [$startColumn, $startRow] = Coordinate::splitCellRef($parts[0]);
                [$endColumn, $endRow] = Coordinate::splitCellRef($parts[1]);
            } catch (\Throwable) {
                continue;
            }
            $ranges[] = [
                'start_row' => min($startRow, $endRow),
                'end_row' => max($startRow, $endRow),
                'start_column' => min($startColumn, $endColumn),
                'end_column' => max($startColumn, $endColumn),
                'ref' => $ref,
            ];
            if ($maxRanges > 0 && count($ranges) > $maxRanges) {
                throw MnbExcelException::withCode('XLSX merged-cell range limit exceeded.', ErrorCode::FILE_READ_FAILED, ['max_merged_ranges' => $maxRanges]);
            }
        }
        return new self($ranges);
    }

    public function active(): bool
    {
        return $this->ranges !== [];
    }

    /** @return list<array{start_row:int,end_row:int,start_column:int,end_column:int,ref:string}> */
    public function ranges(): array
    {
        return $this->ranges;
    }

    /**
     * @param list<mixed> $row
     * @param array<string,mixed> $anchors
     * @return list<mixed>
     */
    public function expandRow(int $rowNumber, array $row, array &$anchors): array
    {
        foreach ($this->ranges as $range) {
            if ($rowNumber < $range['start_row'] || $rowNumber > $range['end_row']) {
                continue;
            }
            $key = $range['ref'];
            if ($rowNumber === $range['start_row']) {
                $anchors[$key] = $row[$range['start_column'] - 1] ?? null;
            }
            $value = $anchors[$key] ?? null;
            for ($column = $range['start_column']; $column <= $range['end_column']; $column++) {
                $row[$column - 1] = $value;
            }
        }
        if ($row === []) {
            return [];
        }
        ksort($row);
        $max = max(array_keys($row));
        $dense = [];
        for ($i = 0; $i <= $max; $i++) {
            $dense[] = $row[$i] ?? null;
        }
        return $dense;
    }
}
