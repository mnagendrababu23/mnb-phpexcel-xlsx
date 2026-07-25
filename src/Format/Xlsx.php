<?php

declare(strict_types=1);

namespace Mnb\PHPExcel\Format;

use Mnb\PHPExcel\Core\WorkbookFactory;
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
