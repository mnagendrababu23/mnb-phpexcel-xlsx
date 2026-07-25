<?php

declare(strict_types=1);

namespace Mnb\PHPExcel\Large;

/**
 * Converts XLSX preflight metrics into a safe import strategy.
 *
 * This class intentionally does not open or read workbooks. It is pure decision logic,
 * so applications can unit-test threshold behaviour and override server constraints.
 */
final class ImportMethodAdvisor
{
    public const METHOD_NORMAL_READ = 'normal_read';
    public const METHOD_NORMAL_READ_WITH_WARNING = 'normal_read_with_warning';
    public const METHOD_STREAMING_READ = 'streaming_read';
    public const METHOD_STREAMING_CHUNK_IMPORT = 'streaming_chunk_import';
    public const METHOD_CLI_QUEUE_RECOMMENDED = 'cli_queue_recommended';
    public const METHOD_CSV_ZIP_OR_SPLIT_RECOMMENDED = 'csv_zip_or_split_recommended';
    public const METHOD_UNSUPPORTED_OR_HIGH_RISK = 'unsupported_or_high_risk';

    private const EXCEL_MAX_ROWS = 1048576;
    private const EXCEL_MAX_COLUMNS = 16384;

    /**
     * @param array<string,mixed> $profile A preflight profile from LargeExcelPreflightAnalyzer::analyze().
     * @param array<string,mixed> $serverOptions server, memory_limit, max_execution_time, allow_http_large_import, prefer_cli, etc.
     * @return array<string,mixed>
     */
    public function recommendFromProfile(array $profile, array $serverOptions = []): array
    {
        $server = $this->resolveServerProfile($serverOptions);
        $largestSheet = $this->largestSheet($profile);
        $rows = (int) ($profile['total_rows'] ?? $profile['total_estimated_rows'] ?? ($largestSheet['rows'] ?? 0));
        $columns = (int) ($profile['max_columns'] ?? ($largestSheet['columns'] ?? 0));
        $cells = (int) ($profile['total_cells'] ?? $profile['total_estimated_cells'] ?? max(0, $rows * max(1, $columns)));
        $fileSizeMb = (float) ($profile['file_size_mb'] ?? 0.0);
        $uncompressedMb = (float) ($profile['total_uncompressed_sheet_xml_mb'] ?? 0.0);
        $sheetCount = (int) ($profile['sheet_count'] ?? count((array) ($profile['sheets'] ?? [])));
        $featureScore = $this->featureScore((array) ($profile['features'] ?? []));
        $level = $this->level($rows, $cells, $fileSizeMb, $uncompressedMb, $columns);

        $reasons = [];
        $warnings = [];
        $method = self::METHOD_NORMAL_READ;
        $runtime = 'http_or_cli';
        $normalAllowed = true;

        if ($rows > self::EXCEL_MAX_ROWS || $columns > self::EXCEL_MAX_COLUMNS) {
            $method = self::METHOD_CSV_ZIP_OR_SPLIT_RECOMMENDED;
            $runtime = 'cli_or_queue_only';
            $normalAllowed = false;
            $reasons[] = 'Worksheet exceeds Microsoft Excel sheet limits; split sheets or import as CSV/DB parts.';
        } elseif ($level === 'tiny' || $level === 'small') {
            $method = self::METHOD_NORMAL_READ;
            $runtime = 'http_or_cli';
            $reasons[] = 'Workbook is within small-file thresholds; current rich reader is safe.';
        } elseif ($level === 'normal') {
            if ($server['memory_mb'] >= 512 && $featureScore <= 2) {
                $method = self::METHOD_NORMAL_READ_WITH_WARNING;
                $runtime = 'http_or_cli_with_timeout_guard';
                $reasons[] = 'Workbook is normal-sized and simple, but streaming is safer on low-memory servers.';
            } else {
                $method = self::METHOD_STREAMING_CHUNK_IMPORT;
                $runtime = 'http_possible_cli_preferred';
                $normalAllowed = false;
                $reasons[] = 'Workbook is normal-sized but server/features make full-array loading risky.';
            }
        } elseif ($level === 'medium') {
            $method = self::METHOD_STREAMING_CHUNK_IMPORT;
            $runtime = $server['max_execution_time'] > 0 && $server['max_execution_time'] <= 120 ? 'cli_or_queue_recommended' : 'http_possible_cli_preferred';
            $normalAllowed = false;
            $reasons[] = 'Medium workbook should be processed in streamed chunks to avoid PHP array memory growth.';
        } elseif ($level === 'large') {
            $method = self::METHOD_STREAMING_CHUNK_IMPORT;
            $runtime = 'cli_or_queue_recommended';
            $normalAllowed = false;
            $reasons[] = 'Large workbook requires streaming chunk import; current normal reader should not be used.';
        } else { // very_large / extreme
            $method = self::METHOD_CLI_QUEUE_RECOMMENDED;
            $runtime = 'cli_or_queue_only';
            $normalAllowed = false;
            $reasons[] = 'Very large workbook should run outside normal HTTP requests with checkpoint/progress support.';
        }

        if ($server['profile'] === 'shared_128' || $server['profile'] === 'shared_256') {
            if (in_array($method, [self::METHOD_STREAMING_CHUNK_IMPORT, self::METHOD_CLI_QUEUE_RECOMMENDED], true)) {
                $warnings[] = 'Shared hosting should use very small chunks and preferably CLI/queue to avoid timeout.';
            }
        }

        if ($featureScore >= 4) {
            $warnings[] = 'Workbook has rich/advanced Excel features; import should stream values and treat objects as metadata/warnings.';
        }

        if ($sheetCount > 5 && $level !== 'tiny' && $level !== 'small') {
            $warnings[] = 'Many worksheets detected; import sheet-by-sheet rather than all sheets in one request.';
        }

        $chunkSize = $this->chunkSize($level, $server['profile'], $server['memory_mb'], $featureScore, $columns);
        $timeBudget = $this->timeBudget($server, $method);

        return [
            'status' => 'ok',
            'method' => $method,
            'level' => $level,
            'runtime' => $runtime,
            'normal_reader_allowed' => $normalAllowed,
            'recommended_chunk_size' => $chunkSize,
            'recommended_time_budget_seconds' => $timeBudget,
            'recommended_route' => $this->routeFor($method),
            'server_profile' => $server,
            'metrics' => [
                'rows' => $rows,
                'columns' => $columns,
                'cells' => $cells,
                'file_size_mb' => round($fileSizeMb, 2),
                'uncompressed_sheet_xml_mb' => round($uncompressedMb, 2),
                'sheet_count' => $sheetCount,
                'feature_score' => $featureScore,
            ],
            'reasons' => $reasons,
            'warnings' => $warnings,
            'matrix' => $this->decisionMatrix(),
        ];
    }

    /** @param array<string,mixed> $serverOptions @return array<string,mixed> */
    public function resolveServerProfile(array $serverOptions = []): array
    {
        $memoryLimit = (string) ($serverOptions['memory_limit'] ?? ini_get('memory_limit') ?: '128M');
        $memoryMb = $this->memoryToMb($memoryLimit);
        $maxExecutionTime = isset($serverOptions['max_execution_time'])
            ? (int) $serverOptions['max_execution_time']
            : (int) (ini_get('max_execution_time') ?: 0);

        $explicit = strtolower((string) ($serverOptions['server'] ?? $serverOptions['profile'] ?? 'auto'));
        $profile = match ($explicit) {
            'shared', 'shared_128', 'shared-128' => $memoryMb <= 128 ? 'shared_128' : 'shared_256',
            'shared_256', 'shared-256' => 'shared_256',
            'vps', 'vps_512', 'vps-512' => $memoryMb <= 512 ? 'vps_512' : 'vps_1g',
            'vps_1g', 'vps-1g', 'vps_1024' => 'vps_1g',
            'worker', 'dedicated', 'server', 'worker_2g', 'worker-2g' => 'worker_2g',
            default => $this->autoProfile($memoryMb),
        };

        return [
            'profile' => $profile,
            'memory_limit' => $memoryLimit,
            'memory_mb' => $memoryMb,
            'max_execution_time' => $maxExecutionTime,
            'allow_http_large_import' => (bool) ($serverOptions['allow_http_large_import'] ?? false),
            'prefer_cli' => (bool) ($serverOptions['prefer_cli'] ?? true),
        ];
    }

    /** @return array<string,mixed> */
    public function decisionMatrix(): array
    {
        return [
            'tiny' => ['rows' => '1-2,000', 'cells' => '<=50,000', 'method' => self::METHOD_NORMAL_READ],
            'small' => ['rows' => '2,001-10,000', 'cells' => '<=150,000', 'method' => self::METHOD_NORMAL_READ],
            'normal' => ['rows' => '10,001-50,000', 'cells' => '<=500,000', 'method' => self::METHOD_NORMAL_READ_WITH_WARNING],
            'medium' => ['rows' => '50,001-150,000', 'cells' => '<=1,500,000', 'method' => self::METHOD_STREAMING_CHUNK_IMPORT],
            'large' => ['rows' => '150,001-500,000', 'cells' => '<=5,000,000', 'method' => self::METHOD_STREAMING_CHUNK_IMPORT],
            'very_large' => ['rows' => '500,001-1,048,576', 'cells' => '>5,000,000', 'method' => self::METHOD_CLI_QUEUE_RECOMMENDED],
            'beyond_excel_sheet' => ['rows' => '>1,048,576', 'cells' => 'any', 'method' => self::METHOD_CSV_ZIP_OR_SPLIT_RECOMMENDED],
        ];
    }

    /** @param array<string,mixed> $profile @return array<string,mixed> */
    private function largestSheet(array $profile): array
    {
        $largest = [];
        foreach ((array) ($profile['sheets'] ?? []) as $sheet) {
            if (!is_array($sheet)) {
                continue;
            }
            if ((int) ($sheet['rows'] ?? 0) > (int) ($largest['rows'] ?? 0)) {
                $largest = $sheet;
            }
        }
        return $largest;
    }

    private function level(int $rows, int $cells, float $fileSizeMb, float $uncompressedMb, int $columns): string
    {
        if ($rows > self::EXCEL_MAX_ROWS || $columns > self::EXCEL_MAX_COLUMNS) {
            return 'beyond_excel_sheet';
        }
        if ($rows <= 2000 && $cells <= 50000 && $fileSizeMb <= 5.0) {
            return 'tiny';
        }
        if ($rows <= 10000 && $cells <= 150000 && $fileSizeMb <= 20.0) {
            return 'small';
        }
        if ($rows <= 50000 && $cells <= 500000 && $fileSizeMb <= 50.0 && $uncompressedMb <= 150.0) {
            return 'normal';
        }
        if ($rows <= 150000 && $cells <= 1500000 && $fileSizeMb <= 100.0 && $uncompressedMb <= 350.0) {
            return 'medium';
        }
        if ($rows <= 500000 && $cells <= 5000000 && $fileSizeMb <= 250.0 && $uncompressedMb <= 900.0) {
            return 'large';
        }
        return 'very_large';
    }

    /** @param array<string,mixed> $features */
    private function featureScore(array $features): int
    {
        $score = 0;
        foreach (['formulas', 'comments', 'hyperlinks', 'merged_cells', 'drawings', 'charts', 'pivot_tables', 'external_links', 'macros'] as $key) {
            $value = $features[$key] ?? false;
            if ($value === true || (is_numeric($value) && (int) $value > 0)) {
                $score++;
            }
        }
        return $score;
    }

    private function chunkSize(string $level, string $profile, int $memoryMb, int $featureScore, int $columns): int
    {
        $base = match ($profile) {
            'shared_128' => match ($level) {
                'medium' => 250,
                'large', 'very_large' => 100,
                default => 500,
            },
            'shared_256' => match ($level) {
                'medium' => 500,
                'large', 'very_large' => 250,
                default => 1000,
            },
            'vps_512' => match ($level) {
                'medium' => 1000,
                'large' => 1000,
                'very_large' => 500,
                default => 2000,
            },
            'vps_1g' => match ($level) {
                'medium' => 2000,
                'large' => 2500,
                'very_large' => 1000,
                default => 5000,
            },
            default => match ($level) {
                'medium' => 5000,
                'large' => 5000,
                'very_large' => 2500,
                default => 10000,
            },
        };

        if ($memoryMb <= 128) {
            $base = min($base, 250);
        }
        if ($featureScore >= 4) {
            $base = max(50, (int) floor($base / 2));
        }
        if ($columns > 50) {
            $base = max(50, (int) floor($base / 2));
        }

        return max(25, $base);
    }

    /** @param array<string,mixed> $server */
    private function timeBudget(array $server, string $method): int
    {
        $max = (int) ($server['max_execution_time'] ?? 0);
        if ($max <= 0) {
            return 0;
        }
        if (in_array($method, [self::METHOD_NORMAL_READ, self::METHOD_NORMAL_READ_WITH_WARNING], true)) {
            return max(1, $max - 3);
        }
        return max(1, $max - 8);
    }

    private function routeFor(string $method): string
    {
        return match ($method) {
            self::METHOD_NORMAL_READ => 'MnbExcel::read($path)->toArray()',
            self::METHOD_NORMAL_READ_WITH_WARNING => 'MnbExcel::read($path)->toArray() or MnbExcel::largeRead($path)->chunk(...)',
            self::METHOD_STREAMING_READ => 'MnbExcel::largeRead($path)->eachRow(...)',
            self::METHOD_STREAMING_CHUNK_IMPORT => 'MnbExcel::largeRead($path)->chunk($chunkSize, $callback)',
            self::METHOD_CLI_QUEUE_RECOMMENDED => 'Run MnbExcel::largeRead($path)->chunk(...) from CLI/queue with checkpoint/progress',
            self::METHOD_CSV_ZIP_OR_SPLIT_RECOMMENDED => 'Split sheets/CSV ZIP parts or import directly into database',
            default => 'Manual review required before import',
        };
    }

    private function autoProfile(int $memoryMb): string
    {
        if ($memoryMb <= 128) {
            return 'shared_128';
        }
        if ($memoryMb <= 256) {
            return 'shared_256';
        }
        if ($memoryMb <= 512) {
            return 'vps_512';
        }
        if ($memoryMb <= 1024) {
            return 'vps_1g';
        }
        return 'worker_2g';
    }

    private function memoryToMb(string $value): int
    {
        $value = trim($value);
        if ($value === '' || $value === '-1') {
            return 4096;
        }
        if (!preg_match('/^([0-9.]+)\s*([KMG])?B?$/i', $value, $m)) {
            return 128;
        }
        $amount = (float) $m[1];
        $unit = strtoupper($m[2] ?? 'M');
        return (int) round(match ($unit) {
            'K' => $amount / 1024,
            'G' => $amount * 1024,
            default => $amount,
        });
    }
}
