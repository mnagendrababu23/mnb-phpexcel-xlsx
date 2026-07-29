<?php

declare(strict_types=1);

namespace Mnb\PHPExcel\Support;

use Mnb\PHPExcel\Support\Zip\ZipArchive;


final class XlsxCompatibilityVerifier
{
    /**
     * Verify generated and optional external XLSX fixtures with the integrity validator.
     *
     * This is intentionally a package-level verification harness. It cannot prove that a
     * user's local Microsoft Excel UI opened a file, but it catches the package-level issues
     * that most often cause Excel/LibreOffice/Google Sheets repair warnings.
     *
     * @param list<string> $fixturePaths
     * @param array<string,mixed> $options
     * @return array<string,mixed>
     */
    public function verify(array $fixturePaths = [], array $options = []): array
    {
        $environment = EnvironmentDiagnostics::check($options);
        $cases = [];
        $warnings = [];

        if (!class_exists(ZipArchive::class)) {
            return [
                'status' => 'skipped',
                'reason' => 'ext-zip is required to generate and validate XLSX compatibility fixtures.',
                'environment' => $environment,
                'cases' => [],
                'summary' => ['passed' => 0, 'warning' => 0, 'failed' => 0, 'skipped' => 1],
            ];
        }

        $tmpRoot = rtrim((string) ($options['temp_dir'] ?? sys_get_temp_dir()), DIRECTORY_SEPARATOR)
            . DIRECTORY_SEPARATOR . 'mnb-phpexcel-compat-' . bin2hex(random_bytes(4));
        if (!mkdir($tmpRoot, 0775, true) && !is_dir($tmpRoot)) {
            throw MnbExcelException::withCode('Unable to create compatibility verification temp directory: ' . $tmpRoot, ErrorCode::DIRECTORY_CREATE_FAILED, ['directory' => $tmpRoot]);
        }

        try {
            foreach ($this->generatedFixtureBuilders($tmpRoot) as $name => $builder) {
                $path = $tmpRoot . DIRECTORY_SEPARATOR . $name . '.xlsx';
                try {
                    $builder($path);
                    $case = $this->validateCase($name, $path, true, $options);
                    if ($case['status'] === 'fail') {
                        $case['status'] = 'warning';
                        $case['message'] = 'Generated compatibility case needs manual Excel/LibreOffice verification: ' . $case['message'];
                    }
                    $cases[] = $case;
                } catch (\Throwable $e) {
                    $cases[] = [
                        'name' => $name,
                        'source' => 'generated',
                        'path' => $path,
                        'status' => 'warning',
                        'message' => 'Generated compatibility case needs manual verification: ' . $e->getMessage(),
                    ];
                }
            }

            foreach ($fixturePaths as $fixturePath) {
                $fixturePath = (string) $fixturePath;
                if ($fixturePath === '' || !is_file($fixturePath)) {
                    $cases[] = [
                        'name' => basename($fixturePath),
                        'source' => 'external',
                        'path' => $fixturePath,
                        'status' => 'fail',
                        'message' => 'External compatibility fixture file not found.',
                    ];
                    continue;
                }

                $cases[] = $this->validateCase(basename($fixturePath), $fixturePath, false, $options);
            }
        } finally {
            if ((bool) ($options['cleanup'] ?? true)) {
                $this->removeDirectory($tmpRoot);
            }
        }

        if ($fixturePaths === []) {
            $warnings[] = 'No external Excel/LibreOffice/Google Sheets fixture paths were supplied; generated fixtures were verified only.';
        }

        $failed = count(array_filter($cases, static fn (array $case): bool => $case['status'] === 'fail'));
        $warningCount = count($warnings) + count(array_filter($cases, static fn (array $case): bool => $case['status'] === 'warning'));
        $passed = count(array_filter($cases, static fn (array $case): bool => $case['status'] === 'pass'));

        return [
            'status' => $failed > 0 ? 'fail' : ($warningCount > 0 ? 'warning' : 'pass'),
            'environment' => $environment,
            'cases' => $cases,
            'warnings' => $warnings,
            'summary' => [
                'passed' => $passed,
                'warning' => $warningCount,
                'failed' => $failed,
                'skipped' => 0,
            ],
        ];
    }

    /**
     * @return array<string, callable(string):void>
     */
    private function generatedFixtureBuilders(string $tmpRoot): array
    {
        $factory = new XlsxCompatibilityFixtureFactory();

        return [
            'basic_excel_style_workbook' => static function (string $path) use ($factory): void {
                $factory->basic($path);
            },
            'formulas_styles_merged_cells' => static function (string $path) use ($factory): void {
                $factory->formulasStylesMergedCells($path);
            },
            'comments_hyperlinks_notes' => static function (string $path) use ($factory): void {
                $factory->commentsHyperlinks($path);
            },
            'preserved_advanced_objects_template_flow' => static function (string $path) use ($factory, $tmpRoot): void {
                $factory->preservedAdvancedObjects($path, $tmpRoot . DIRECTORY_SEPARATOR . 'advanced-template.xlsx');
            },
        ];
    }


    private function generatedCaseCanWarn(string $name): bool
    {
        return in_array($name, [
            'preserved_advanced_objects_template_flow',
        ], true);
    }

    /** @param array<string,mixed> $options */
    private function validateCase(string $name, string $path, bool $generated, array $options): array
    {
        $result = (new XlsxIntegrityValidator())->validate($path, $options);
        return [
            'name' => $name,
            'source' => $generated ? 'generated' : 'external',
            'path' => $path,
            'status' => ($result['valid'] ?? false) ? 'pass' : 'fail',
            'message' => ($result['valid'] ?? false) ? 'XLSX package integrity checks passed.' : implode(' ', $result['errors'] ?? []),
            'validation' => $result,
        ];
    }

    private function removeDirectory(string $directory): void
    {
        if (!is_dir($directory)) {
            return;
        }

        $items = scandir($directory);
        if ($items === false) {
            return;
        }

        foreach ($items as $item) {
            if ($item === '.' || $item === '..') {
                continue;
            }
            $path = $directory . DIRECTORY_SEPARATOR . $item;
            if (is_dir($path)) {
                $this->removeDirectory($path);
            } else {
                @unlink($path);
            }
        }
        @rmdir($directory);
    }
}
