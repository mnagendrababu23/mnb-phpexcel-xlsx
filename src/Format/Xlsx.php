<?php

declare(strict_types=1);

namespace Mnb\PHPExcel\Format;

use Mnb\PHPExcel\Core\RichText;
use Mnb\PHPExcel\Core\WorkbookBuilder;
use Mnb\PHPExcel\Core\WorkbookFactory;
use Mnb\PHPExcel\Reader\State\CellSnapshot;
use Mnb\PHPExcel\Reader\Options\ReaderOptions;
use Mnb\PHPExcel\Reader\ReadSession;
use Mnb\PHPExcel\Reader\XlsxReader;
use Mnb\PHPExcel\Writer\XlsxWriter;

final class Xlsx
{
    /** @param array<string,mixed>|ReaderOptions $options */
    public static function read(string $path, array|ReaderOptions $options = []): ReadSession
    {
        return new ReadSession($path, new XlsxReader(), $options);
    }

    public static function cell(string $path, string $cell, int|string $sheet = 1, array|ReaderOptions $options = []): mixed
    {
        return self::read($path, $options)->sheet($sheet)->cell($cell);
    }

    public static function cellDetails(string $path, string $cell, int|string $sheet = 1, array|ReaderOptions $options = []): CellSnapshot
    {
        return self::read($path, $options)->sheet($sheet)->cellDetails($cell);
    }

    /** @return array<int,array<int,mixed>> */
    public static function rangeValues(string $path, string $range, int|string $sheet = 1, array|ReaderOptions $options = []): array
    {
        return self::read($path, $options)->sheet($sheet)->rangeValues($range);
    }

    public static function richText(string $path, string $cell, int|string $sheet = 1): ?RichText
    {
        return self::read($path)->sheet($sheet)->richText($cell);
    }

    /** @return list<array<string,mixed>> */
    public static function images(string $path, int|string $sheet = 1, bool $includeBytes = false): array
    {
        return self::read($path)->sheet($sheet)->images($includeBytes);
    }

    /** Create a styled, validated XLSX import template. @param array<int|string,string|array<string,mixed>> $columns */
    public static function writeImportTemplate(array $columns, string $path, array $options = []): void
    {
        WorkbookBuilder::importTemplate($columns, $options)->toXlsx($path);
    }

    /** @param iterable<array<int|string,mixed>|mixed> $rows @param array<string,mixed> $options */
    public static function write(iterable $rows, string $path, array $options = []): void
    {
        $workbook = WorkbookFactory::workbook(
            $rows,
            (string) ($options['sheet_name'] ?? 'Sheet1'),
            (bool) ($options['with_header'] ?? false)
        );
        (new XlsxWriter())->write($workbook, $path);
    }
}
