<?php

declare(strict_types=1);

namespace Mnb\PHPExcel\Support;

final class CompatibilityFixtureSuite
{
    /** @return array<string,mixed> */
    public static function requirements(): array
    {
        return [
            'status' => 'ready',
            'required_fixture_groups' => [
                'microsoft_excel' => ['simple', 'formulas', 'styles', 'comments_hyperlinks', 'merged_hidden_multisheet'],
                'libreoffice_calc' => ['simple', 'formulas', 'styles', 'comments_hyperlinks', 'merged_hidden_multisheet'],
                'google_sheets_export' => ['simple', 'formulas', 'styles', 'comments_hyperlinks'],
                'wps_office' => ['simple', 'formulas', 'styles'],
            ],
            'checks' => [
                'XLSX integrity validation',
                'preflight sheet/row/column detection',
                'normal reader small fixture accuracy',
                'large streaming reader values/chunks',
                'comments/hyperlinks metadata extraction',
                'styles/date/numeric conversion sanity',
                'manual open without repair warning',
            ],
            'naming' => [
                'docs/fixtures/microsoft-excel/simple.xlsx',
                'docs/fixtures/libreoffice-calc/formulas.xlsx',
                'docs/fixtures/google-sheets-export/comments-hyperlinks.xlsx',
                'docs/fixtures/wps-office/styles.xlsx',
            ],
        ];
    }

    /** @return array<string,mixed> */
    public static function verify(string $fixtureDir, array $options = []): array
    {
        $strict = (bool) ($options['strict'] ?? false);
        $files = self::findXlsxFiles($fixtureDir);
        $results = [];
        $summary = ['passed' => 0, 'warning' => 0, 'failed' => 0, 'skipped' => 0];

        if ($files === []) {
            $summary[$strict ? 'failed' : 'warning']++;
            return [
                'status' => $strict ? 'failed' : 'warning',
                'fixture_dir' => $fixtureDir,
                'summary' => $summary,
                'results' => [],
                'message' => 'No real XLSX fixture files were found. Add Excel/LibreOffice/Google Sheets/WPS fixture files before publishing compatibility numbers.',
                'requirements' => self::requirements(),
            ];
        }

        $validator = new XlsxIntegrityValidator();
        $verifier = new XlsxCompatibilityVerifier();
        foreach ($files as $file) {
            try {
                $integrity = $validator->validate($file, ['strict_xml' => (bool) ($options['strict_xml'] ?? false)]);
                $compatibility = $verifier->verify([$file], ['include_generated' => false]);
                $status = (($integrity['valid'] ?? false) && (int) ($compatibility['summary']['failed'] ?? 0) === 0) ? 'passed' : 'warning';
                $summary[$status === 'passed' ? 'passed' : 'warning']++;
                $results[] = [
                    'path' => $file,
                    'status' => $status,
                    'integrity_status' => $integrity['status'] ?? 'unknown',
                    'compatibility_status' => $compatibility['status'] ?? 'unknown',
                    'manual_excel_open_required' => true,
                    'notes' => $status === 'passed' ? [] : ['Review warnings and manually open the file in Excel/LibreOffice.'],
                ];
            } catch (\Throwable $e) {
                $summary['failed']++;
                $results[] = [
                    'path' => $file,
                    'status' => 'failed',
                    'error' => $e->getMessage(),
                    'manual_excel_open_required' => true,
                ];
            }
        }

        return [
            'status' => $summary['failed'] > 0 ? 'failed' : ($summary['warning'] > 0 ? 'warning' : 'passed'),
            'fixture_dir' => $fixtureDir,
            'summary' => $summary,
            'results' => $results,
            'requirements' => self::requirements(),
        ];
    }

    /** @return list<string> */
    private static function findXlsxFiles(string $dir): array
    {
        if (!is_dir($dir)) {
            return [];
        }
        $files = [];
        $iterator = new \RecursiveIteratorIterator(new \RecursiveDirectoryIterator($dir, \FilesystemIterator::SKIP_DOTS));
        foreach ($iterator as $file) {
            if ($file instanceof \SplFileInfo && $file->isFile() && strtolower($file->getExtension()) === 'xlsx') {
                $files[] = $file->getPathname();
            }
        }
        sort($files);
        return $files;
    }
}
