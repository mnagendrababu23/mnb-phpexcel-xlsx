<?php

declare(strict_types=1);

namespace Mnb\PHPExcel\Large;

use Mnb\PHPExcel\Reader\SharedStrings\SharedStringProviderInterface;

use Mnb\PHPExcel\Support\ErrorCode;
use Mnb\PHPExcel\Support\MnbExcelException;
use PDO;
use Throwable;
use XMLReader;

/**
 * Shared strings cache for streaming XLSX import.
 *
 * Small tables are kept in memory. Large tables use a temporary SQLite database when available,
 * so the worksheet reader does not need to load a huge sharedStrings.xml into RAM.
 */
final class LargeSharedStringCache implements SharedStringProviderInterface
{
    /** @var list<string> */
    private array $memory = [];
    private ?PDO $pdo = null;
    private ?string $sqlitePath = null;
    private ?\PDOStatement $insertStatement = null;
    private string $mode = 'none';
    private int $count = 0;

    /** @param array<string,mixed> $options */
    public static function fromXlsx(string $realPath, array $options = []): self
    {
        $cache = new self();
        $tempSharedStrings = self::extractEntryToTempFile(
            $realPath,
            'xl/sharedStrings.xml',
            (string) ($options['temp_dir'] ?? sys_get_temp_dir())
        );

        if ($tempSharedStrings === null) {
            return $cache;
        }

        try {
            $reader = self::openXmlFile($tempSharedStrings, 'XLSX shared string table');

            $declaredCount = 0;
            while ($reader->read()) {
                if ($reader->nodeType === XMLReader::ELEMENT && $reader->name === 'sst') {
                    $declaredCount = (int) ($reader->getAttribute('uniqueCount') ?: $reader->getAttribute('count') ?: 0);
                    break;
                }
            }
            $reader->close();

            $maxMemoryStrings = (int) ($options['max_shared_strings_in_memory'] ?? 100000);
            $maxMemoryXmlBytes = (int) ($options['max_shared_strings_xml_memory_bytes'] ?? 67108864);
            $sharedStringsXmlBytes = (int) (filesize($tempSharedStrings) ?: 0);
            $preferSqlite = (bool) ($options['prefer_sqlite_shared_strings'] ?? true);
            $requiresDiskCache = $declaredCount > $maxMemoryStrings || $sharedStringsXmlBytes > $maxMemoryXmlBytes;
            $useSqlite = $preferSqlite && $requiresDiskCache && self::sqliteAvailable();

            if ($requiresDiskCache && !$useSqlite) {
                throw MnbExcelException::withCode(
                    'Shared string table is too large for memory mode and pdo_sqlite is not available. Enable pdo_sqlite or increase max_shared_strings_in_memory/max_shared_strings_xml_memory_bytes intentionally.',
                    ErrorCode::EXTENSION_MISSING,
                    [
                        'declared_shared_strings' => $declaredCount,
                        'shared_strings_xml_bytes' => $sharedStringsXmlBytes,
                        'max_shared_strings_in_memory' => $maxMemoryStrings,
                        'max_shared_strings_xml_memory_bytes' => $maxMemoryXmlBytes,
                    ],
                    null,
                    'The Excel shared string table is too large for this server configuration.'
                );
            }

            if ($useSqlite) {
                $cache->initializeSqlite((string) ($options['temp_dir'] ?? sys_get_temp_dir()));
            } else {
                $cache->mode = 'memory';
            }

            $reader = self::openXmlFile($tempSharedStrings, 'XLSX shared string table');
            $index = 0;
            if ($cache->pdo !== null) {
                $cache->pdo->beginTransaction();
            }
            while ($reader->read()) {
                if ($reader->nodeType !== XMLReader::ELEMENT || $reader->name !== 'si') {
                    continue;
                }
                $text = self::textFromRichXml($reader->readOuterXml());
                $cache->put($index, $text);
                $index++;
            }
            if ($cache->pdo !== null) {
                $cache->pdo->commit();
            }
            $reader->close();
            $cache->count = $index;

            return $cache;
        } finally {
            if (is_file($tempSharedStrings)) {
                @unlink($tempSharedStrings);
            }
        }
    }

    public function get(int $index): string
    {
        if ($this->mode === 'memory') {
            return $this->memory[$index] ?? '';
        }
        if ($this->pdo !== null) {
            $stmt = $this->pdo->prepare('SELECT value FROM shared_strings WHERE id = :id');
            $stmt->execute([':id' => $index]);
            $value = $stmt->fetchColumn();
            return is_string($value) ? $value : '';
        }
        return '';
    }

    public function mode(): string
    {
        return $this->mode;
    }

    public function count(): int
    {
        return $this->count;
    }

    public function close(): void
    {
        $this->insertStatement = null;
        $this->pdo = null;
        if ($this->sqlitePath !== null && is_file($this->sqlitePath)) {
            @unlink($this->sqlitePath);
        }
        $this->sqlitePath = null;
    }

    public function __destruct()
    {
        $this->close();
    }


    private static function extractEntryToTempFile(string $realPath, string $entry, string $tempDir): ?string
    {
        $zip = new \ZipArchive();
        $openResult = $zip->open($realPath);
        if ($openResult !== true) {
            throw MnbExcelException::withCode(
                'Unable to open XLSX package while checking shared strings.',
                ErrorCode::XLSX_INVALID,
                ['zip_status' => $openResult],
                null,
                'The Excel file could not be opened safely.'
            );
        }

        if ($zip->statName($entry) === false) {
            $zip->close();
            return null;
        }

        if (!is_dir($tempDir) || !is_writable($tempDir)) {
            $tempDir = sys_get_temp_dir();
        }
        $tempPath = tempnam($tempDir, 'mnb_xlsx_sst_xml_');
        if ($tempPath === false) {
            $zip->close();
            throw MnbExcelException::withCode('Unable to create temporary shared string XML file.', ErrorCode::FILE_WRITE_FAILED);
        }

        $in = $zip->getStream($entry);
        if (!is_resource($in)) {
            $zip->close();
            @unlink($tempPath);
            throw MnbExcelException::withCode(
                'Unable to open XLSX shared string entry stream.',
                ErrorCode::XLSX_INVALID,
                ['entry' => $entry],
                null,
                'The Excel shared string table could not be read safely.'
            );
        }

        $out = fopen($tempPath, 'wb');
        if (!is_resource($out)) {
            fclose($in);
            $zip->close();
            @unlink($tempPath);
            throw MnbExcelException::withCode('Unable to write temporary shared string XML file.', ErrorCode::FILE_WRITE_FAILED);
        }

        try {
            while (!feof($in)) {
                $chunk = fread($in, 1024 * 1024);
                if ($chunk === false) {
                    throw MnbExcelException::withCode(
                        'Unable to read XLSX shared string entry stream.',
                        ErrorCode::XLSX_INVALID,
                        ['entry' => $entry],
                        null,
                        'The Excel shared string table could not be read safely.'
                    );
                }
                if ($chunk !== '' && fwrite($out, $chunk) === false) {
                    throw MnbExcelException::withCode('Unable to write temporary shared string XML file.', ErrorCode::FILE_WRITE_FAILED);
                }
            }
        } finally {
            fclose($in);
            fclose($out);
            $zip->close();
        }

        return $tempPath;
    }

    private static function openXmlFile(string $path, string $label): XMLReader
    {
        $reader = new XMLReader();
        if (!@$reader->open($path, null, LIBXML_NONET | LIBXML_COMPACT)) {
            throw MnbExcelException::withCode(
                'Unable to open ' . $label . ' for streaming import.',
                ErrorCode::XLSX_INVALID,
                ['path' => $path],
                null,
                'The Excel shared string table could not be read safely.'
            );
        }
        return $reader;
    }

    private static function sqliteAvailable(): bool
    {
        return class_exists(PDO::class) && in_array('sqlite', PDO::getAvailableDrivers(), true);
    }

    private function initializeSqlite(string $tempDir): void
    {
        if (!is_dir($tempDir) || !is_writable($tempDir)) {
            $tempDir = sys_get_temp_dir();
        }
        $path = tempnam($tempDir, 'mnb_xlsx_sst_');
        if ($path === false) {
            throw MnbExcelException::withCode('Unable to create shared string cache file.', ErrorCode::FILE_WRITE_FAILED);
        }
        $this->sqlitePath = $path;
        try {
            $this->pdo = new PDO('sqlite:' . $path);
            $this->pdo->setAttribute(PDO::ATTR_ERRMODE, PDO::ERRMODE_EXCEPTION);
            $this->pdo->exec('CREATE TABLE shared_strings (id INTEGER PRIMARY KEY, value TEXT NOT NULL)');
            $this->pdo->exec('PRAGMA synchronous = OFF');
            $this->pdo->exec('PRAGMA journal_mode = MEMORY');
            $this->mode = 'sqlite';
        } catch (Throwable $e) {
            @unlink($path);
            throw MnbExcelException::withCode('Unable to initialize SQLite shared string cache: ' . $e->getMessage(), ErrorCode::FILE_WRITE_FAILED, [], $e);
        }
    }

    private function put(int $index, string $value): void
    {
        if ($this->pdo !== null) {
            if ($this->insertStatement === null) {
                $this->insertStatement = $this->pdo->prepare('INSERT INTO shared_strings (id, value) VALUES (:id, :value)');
            }
            $this->insertStatement->execute([':id' => $index, ':value' => $value]);
            return;
        }
        $this->memory[$index] = $value;
    }

    private static function textFromRichXml(string $xml): string
    {
        if ($xml === '') {
            return '';
        }
        preg_match_all('/<t\b[^>]*>(.*?)<\/t>/su', $xml, $matches);
        if (($matches[1] ?? []) !== []) {
            return html_entity_decode(implode('', $matches[1]), ENT_QUOTES | ENT_XML1, 'UTF-8');
        }
        return html_entity_decode(strip_tags($xml), ENT_QUOTES | ENT_XML1, 'UTF-8');
    }
}
