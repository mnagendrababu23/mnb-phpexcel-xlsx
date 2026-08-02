<?php

declare(strict_types=1);

namespace Mnb\PHPExcel\Format;

use Mnb\PHPExcel\Core\RichText;
use Mnb\PHPExcel\Template\XlsxImportTemplateFactory;
use Mnb\PHPExcel\Core\WorkbookFactory;
use Mnb\PHPExcel\Reader\State\CellSnapshot;
use Mnb\PHPExcel\Reader\Options\ReaderOptions;
use Mnb\PHPExcel\Reader\ReadSession;
use Mnb\PHPExcel\Reader\XlsxReader;
use Mnb\PHPExcel\Reader\XlsxQuickInfo;
use Mnb\PHPExcel\Reader\XlsxMetadataReader;
use Mnb\PHPExcel\Writer\XlsxWriter;
use Mnb\PHPExcel\Writer\XlsxMetadataWriter;
use Mnb\PHPExcel\Security\XlsxEncryption;

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

    public static function richText(string $path, string $cell, int|string $sheet = 1, array|ReaderOptions $options = []): ?RichText
    {
        return self::read($path, $options)->sheet($sheet)->richText($cell);
    }

    /** @return list<array<string,mixed>> */
    public static function images(string $path, int|string $sheet = 1, bool $includeBytes = false, array|ReaderOptions $options = []): array
    {
        return self::read($path, $options)->sheet($sheet)->images($includeBytes);
    }

    /** @param array<string,mixed> $options @return array<string,mixed> */
    public static function metaInfo(string $path, array $options = []): array
    {
        return (new XlsxMetadataReader())->metaInfo($path, $options);
    }

    /** @param array<string,mixed> $changes @param array<string,mixed> $options */
    public static function updateMetaInfo(string $source, string $destination, array $changes, array $options = []): void
    {
        (new XlsxMetadataWriter())->updateMetaInfo($source, $destination, $changes, $options);
    }

    /** @param array<string,mixed> $options */
    public static function removePersonalInfo(string $source, string $destination, array $options = []): void
    {
        (new XlsxMetadataWriter())->removePersonalInfo($source, $destination, $options);
    }

    /** @param array<string,mixed> $options @return array<string,mixed> */
    public static function fileInfo(string $path, array $options = []): array
    {
        return (new XlsxQuickInfo())->fileInfo($path, $options);
    }

    /** @param array<string,mixed> $options @return list<array<string,mixed>> */
    public static function sheetsInfo(string $path, array $options = []): array
    {
        return (new XlsxQuickInfo())->sheetsInfo($path, $options);
    }

    /** @param array<string,mixed> $options @return array<string,mixed> */
    public static function sheetInfo(string $path, int|string $sheet = 1, array $options = []): array
    {
        return (new XlsxQuickInfo())->sheetInfo($path, $sheet, $options);
    }

    /** @param array<string,mixed> $options */
    public static function rowCount(string $path, int|string $sheet = 1, array $options = []): int
    {
        return (new XlsxQuickInfo())->rowCount($path, $sheet, $options);
    }

    /** @param array<string,mixed> $options @return array<string,int> */
    public static function rowCounts(string $path, array $options = []): array
    {
        return (new XlsxQuickInfo())->rowCounts($path, $options);
    }

    public static function isEncrypted(string $path): bool
    {
        return (new XlsxEncryption())->isEncryptedFile($path);
    }

    /** Return agile, standard, unknown, or null when the file is not encrypted. */
    public static function encryptionMode(string $path): ?string
    {
        return (new XlsxEncryption())->encryptionMode($path);
    }

    /** @param array<string,mixed> $options */
    public static function encryptFile(string $source, string $destination, string $password, array $options = []): void
    {
        (new XlsxEncryption())->encryptFile($source, $destination, $password, $options);
    }

    /** @param array<string,mixed> $options */
    public static function decryptFile(string $source, string $destination, string $password, array $options = []): void
    {
        (new XlsxEncryption())->decryptFile($source, $destination, $password, $options);
    }

    /** @return array<string,mixed> */
    public static function protection(string $path, int|string $sheet = 1, array|ReaderOptions $options = []): array
    {
        return self::read($path, $options)->sheet($sheet)->protection();
    }

    /** Create a styled, validated XLSX import template. @param array<int|string,string|array<string,mixed>> $columns */
    public static function writeImportTemplate(array $columns, string $path, array $options = []): void
    {
        (new XlsxWriter())->write((new XlsxImportTemplateFactory())->create($columns, $options), $path);
    }

    /** @param iterable<array<int|string,mixed>|mixed> $rows @param array<string,mixed> $options */
    public static function write(iterable $rows, string $path, array $options = []): void
    {
        $workbook = WorkbookFactory::workbook(
            $rows,
            (string) ($options['sheet_name'] ?? 'Sheet1'),
            (bool) ($options['with_header'] ?? false)
        );
        $password = (string) ($options['password'] ?? $options['encryption_password'] ?? '');
        if ($password !== '') {
            $encryptionOptions = (array) ($options['encryption_options'] ?? []);
            foreach (['mode', 'encryption_mode', 'compatibility_mode', 'spin_count'] as $key) {
                if (array_key_exists($key, $options)) {
                    $encryptionOptions[$key] = $options[$key];
                }
            }
            $workbook->metadata['_mnb_xlsx_encryption'] = array_replace($encryptionOptions, ['password' => $password]);
        }
        $protectionPassword = (string) ($options['protection_password'] ?? '');
        if ($protectionPassword !== '') {
            if ((bool) ($options['protect_workbook'] ?? true)) {
                $workbook->metadata['_mnb_workbook_protection'] = array_replace((array) ($options['workbook_protection'] ?? []), ['password' => $protectionPassword]);
            }
            if ((bool) ($options['protect_sheets'] ?? true)) {
                $workbook->metadata['_mnb_sheet_protection'] = ['*' => array_replace((array) ($options['sheet_protection'] ?? []), ['password' => $protectionPassword])];
            }
        }
        (new XlsxWriter())->write($workbook, $path);
    }
}
