<?php

declare(strict_types=1);

namespace Mnb\PHPExcel\Application;

use Mnb\PHPExcel\Support\Zip\ZipArchive;

final class UploadSafetyValidator
{
    /** @param array<string,mixed>|string $file @param array<string,mixed> $options @return array<string,mixed> */
    public static function validate(array|string $file, array $options = []): array
    {
        $path = is_array($file) ? (string) ($file['tmp_name'] ?? $file['path'] ?? '') : $file;
        $name = is_array($file) ? (string) ($file['name'] ?? basename($path)) : basename($path);
        $reportedSize = is_array($file) && isset($file['size']) ? max(0, (int) $file['size']) : null;
        $actualSize = is_file($path) ? max(0, (int) filesize($path)) : 0;
        $size = $actualSize > 0 ? $actualSize : ($reportedSize ?? 0);
        $allowed = array_values(array_unique(array_map('strtolower', (array) ($options['allowed_extensions'] ?? ['xlsx', 'xls', 'ods', 'csv', 'tsv', 'json', 'xml']))));
        $maxBytes = max(0, (int) (($options['max_size_mb'] ?? 100) * 1024 * 1024));
        $extension = strtolower(pathinfo($name !== '' ? $name : $path, PATHINFO_EXTENSION));
        $errors = [];
        $warnings = [];
        $features = [];

        if (is_array($file) && isset($file['error']) && (int) $file['error'] !== UPLOAD_ERR_OK) {
            $errors[] = 'Upload failed with PHP upload error code ' . (int) $file['error'] . '.';
        }
        if ($path === '' || !is_file($path)) {
            $errors[] = 'Uploaded file does not exist.';
        } elseif (!is_readable($path)) {
            $errors[] = 'Uploaded file is not readable.';
        }
        if ($path !== '' && is_link($path) && !(bool) ($options['allow_symlinks'] ?? false)) {
            $errors[] = 'Symbolic-link uploads are not allowed.';
        }
        if ($extension === '' || !in_array($extension, $allowed, true)) {
            $errors[] = 'Unsupported upload extension: ' . ($extension ?: 'none') . '.';
        }
        if ($maxBytes > 0 && $size > $maxBytes) {
            $errors[] = 'Uploaded file exceeds max size of ' . (int) ceil($maxBytes / 1024 / 1024) . ' MB.';
        }
        if ($reportedSize !== null && $actualSize > 0 && $reportedSize !== $actualSize) {
            $warnings[] = 'Reported upload size does not match the actual file size.';
        }
        if ($name === '' || str_contains($name, "\0") || str_contains($name, '..') || str_contains($name, '/') || str_contains($name, '\\')) {
            $errors[] = 'Unsafe upload filename.';
        }

        if ($path !== '' && is_file($path) && (bool) ($options['check_mime'] ?? true)) {
            self::inspectMimeType($path, $extension, $options, $features, $warnings, $errors);
        }

        if ($extension === 'xlsx' && $path !== '' && is_file($path)) {
            self::inspectXlsxPackage($path, $options, $features, $warnings, $errors);
        }

        return [
            'status' => $errors === [] ? ($warnings === [] ? 'ok' : 'warning') : 'failed',
            'valid' => $errors === [],
            'path' => $path,
            'name' => $name,
            'extension' => $extension,
            'size_bytes' => $size,
            'reported_size_bytes' => $reportedSize,
            'mime_type' => $features['mime_type'] ?? null,
            'errors' => array_values(array_unique($errors)),
            'warnings' => array_values(array_unique($warnings)),
            'features' => $features,
        ];
    }

    /** @param array<string,mixed> $options @param array<string,mixed> $features @param list<string> $warnings @param list<string> $errors */
    private static function inspectMimeType(string $path, string $extension, array $options, array &$features, array &$warnings, array &$errors): void
    {
        if (!class_exists(\finfo::class)) {
            $warnings[] = 'ext-fileinfo is missing; MIME type could not be checked.';
            return;
        }

        $mime = (new \finfo(FILEINFO_MIME_TYPE))->file($path);
        if (!is_string($mime) || $mime === '') {
            $warnings[] = 'MIME type could not be detected.';
            return;
        }

        $features['mime_type'] = strtolower(trim(explode(';', $mime, 2)[0]));
        $allowedMimes = self::allowedMimeTypes($extension, $options);
        if ($allowedMimes === [] || in_array($features['mime_type'], $allowedMimes, true)) {
            return;
        }

        $message = 'Detected MIME type ' . $features['mime_type'] . ' does not match .' . ($extension ?: 'unknown') . '.';
        if ((bool) ($options['strict_mime'] ?? false)) {
            $errors[] = $message;
        } else {
            $warnings[] = $message;
        }
    }

    /** @param array<string,mixed> $options @return list<string> */
    private static function allowedMimeTypes(string $extension, array $options): array
    {
        $custom = (array) ($options['allowed_mime_types'][$extension] ?? []);
        if ($custom !== []) {
            return array_values(array_unique(array_map('strtolower', array_map('strval', $custom))));
        }

        return match ($extension) {
            'xlsx' => [
                'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                'application/zip',
                'application/x-zip',
                'application/x-zip-compressed',
                'application/octet-stream',
            ],
            'csv', 'tsv' => ['text/csv', 'text/tab-separated-values', 'text/plain', 'application/csv', 'application/vnd.ms-excel', 'application/octet-stream'],
            'xls' => ['application/vnd.ms-excel', 'application/x-ole-storage', 'application/octet-stream'],
            'ods' => ['application/vnd.oasis.opendocument.spreadsheet', 'application/zip', 'application/octet-stream'],
            'json' => ['application/json', 'text/json', 'text/plain', 'application/octet-stream'],
            'xml' => ['application/xml', 'text/xml', 'text/plain', 'application/octet-stream'],
            default => [],
        };
    }

    /** @param array<string,mixed> $options @param array<string,mixed> $features @param list<string> $warnings @param list<string> $errors */
    private static function inspectXlsxPackage(string $path, array $options, array &$features, array &$warnings, array &$errors): void
    {
        if (!class_exists(ZipArchive::class)) {
            $warnings[] = 'ext-zip is missing; XLSX ZIP structure could not be checked.';
            return;
        }

        $zip = new ZipArchive();
        if ($zip->open($path) !== true) {
            $errors[] = 'XLSX upload is not a readable ZIP package.';
            return;
        }

        try {
            foreach (['[Content_Types].xml', '_rels/.rels', 'xl/workbook.xml'] as $entry) {
                if ($zip->locateName($entry) === false) {
                    $errors[] = 'XLSX upload is missing required part: ' . $entry . '.';
                }
            }

            $maxEntries = max(1, (int) ($options['max_zip_entries'] ?? 10000));
            $maxUncompressedBytes = max(1, (int) (($options['max_uncompressed_size_mb'] ?? 1024) * 1024 * 1024));
            $maxEntryBytes = max(1, (int) (($options['max_zip_entry_size_mb'] ?? 256) * 1024 * 1024));
            $maxCompressionRatio = max(1.0, (float) ($options['max_compression_ratio'] ?? 200.0));
            $ratioMinimumBytes = max(0, (int) (($options['compression_ratio_min_size_mb'] ?? 1) * 1024 * 1024));
            $totalUncompressed = 0;
            $totalCompressed = 0;
            $largestEntry = 0;
            $highestRatio = 0.0;
            $unsafePaths = [];
            $suspiciousEntries = [];
            $encryptedEntries = 0;

            if ($zip->numFiles > $maxEntries) {
                $errors[] = 'XLSX ZIP contains too many entries: ' . $zip->numFiles . ' (max ' . $maxEntries . ').';
            }

            for ($i = 0; $i < $zip->numFiles; $i++) {
                $stat = $zip->statIndex($i);
                if (!is_array($stat)) {
                    continue;
                }

                $entryName = (string) ($stat['name'] ?? '');
                $uncompressed = max(0, (int) ($stat['size'] ?? 0));
                $compressed = max(0, (int) ($stat['comp_size'] ?? 0));
                $totalUncompressed += $uncompressed;
                $totalCompressed += $compressed;
                $largestEntry = max($largestEntry, $uncompressed);

                if (self::isUnsafeZipPath($entryName)) {
                    $unsafePaths[] = $entryName;
                }
                if ($uncompressed > $maxEntryBytes) {
                    $suspiciousEntries[] = $entryName . ' (' . $uncompressed . ' bytes)';
                }

                $ratio = $compressed > 0 ? $uncompressed / $compressed : ($uncompressed > 0 ? INF : 0.0);
                $highestRatio = max($highestRatio, is_finite($ratio) ? $ratio : $maxCompressionRatio + 1.0);
                if ($uncompressed >= $ratioMinimumBytes && $ratio > $maxCompressionRatio) {
                    $suspiciousEntries[] = $entryName . ' (compression ratio ' . number_format($ratio, 1) . ':1)';
                }

                if ((int) ($stat['encryption_method'] ?? 0) > 0) {
                    $encryptedEntries++;
                }
            }

            $features['zip_entries'] = $zip->numFiles;
            $features['zip_compressed_bytes'] = $totalCompressed;
            $features['zip_uncompressed_bytes'] = $totalUncompressed;
            $features['zip_largest_entry_bytes'] = $largestEntry;
            $features['zip_highest_compression_ratio'] = round($highestRatio, 2);
            $features['unsafe_zip_paths'] = array_slice(array_values(array_unique($unsafePaths)), 0, 20);
            $features['suspicious_zip_entries'] = array_slice(array_values(array_unique($suspiciousEntries)), 0, 20);
            $features['encrypted_entries'] = $encryptedEntries;
            $features['macros'] = $zip->locateName('xl/vbaProject.bin') !== false;
            $features['external_links'] = self::hasPrefix($zip, 'xl/externalLinks/');
            $features['encrypted'] = $zip->locateName('EncryptedPackage') !== false || $encryptedEntries > 0;

            if ($totalUncompressed > $maxUncompressedBytes) {
                $errors[] = 'XLSX ZIP expands beyond the allowed uncompressed size of ' . (int) ceil($maxUncompressedBytes / 1024 / 1024) . ' MB.';
            }
            if ($unsafePaths !== []) {
                $message = 'Unsafe ZIP entry paths detected: ' . implode(', ', array_slice($unsafePaths, 0, 3)) . '.';
                if ((bool) ($options['reject_unsafe_zip_paths'] ?? true)) {
                    $errors[] = $message;
                } else {
                    $warnings[] = $message;
                }
            }
            if ($suspiciousEntries !== []) {
                $message = 'Potential ZIP-bomb characteristics detected in: ' . implode(', ', array_slice($suspiciousEntries, 0, 3)) . '.';
                if ((bool) ($options['reject_suspicious_archives'] ?? true)) {
                    $errors[] = $message;
                } else {
                    $warnings[] = $message;
                }
            }
            if ($features['macros']) {
                $message = 'Macro-enabled workbook content detected.';
                if ((bool) ($options['reject_macros'] ?? false)) {
                    $errors[] = $message;
                } else {
                    $warnings[] = $message;
                }
            }
            if ($features['external_links']) {
                $message = 'External workbook links detected.';
                if ((bool) ($options['reject_external_links'] ?? false)) {
                    $errors[] = $message;
                } else {
                    $warnings[] = $message;
                }
            }
            if ($features['encrypted']) {
                $message = 'Encrypted workbook package content detected.';
                if ((bool) ($options['reject_encrypted'] ?? true)) {
                    $errors[] = $message;
                } else {
                    $warnings[] = $message;
                }
            }
        } finally {
            $zip->close();
        }
    }

    private static function isUnsafeZipPath(string $name): bool
    {
        $normalized = str_replace('\\', '/', $name);
        return $normalized === ''
            || str_starts_with($normalized, '/')
            || preg_match('/^[A-Za-z]:\//', $normalized) === 1
            || in_array('..', explode('/', $normalized), true)
            || str_contains($normalized, "\0");
    }

    private static function hasPrefix(ZipArchive $zip, string $prefix): bool
    {
        for ($i = 0; $i < $zip->numFiles; $i++) {
            $name = $zip->getNameIndex($i);
            if (is_string($name) && str_starts_with($name, $prefix)) {
                return true;
            }
        }
        return false;
    }
}
