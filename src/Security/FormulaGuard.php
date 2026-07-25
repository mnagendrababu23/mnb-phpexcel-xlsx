<?php

declare(strict_types=1);

namespace Mnb\PHPExcel\Security;

use Mnb\PHPExcel\Support\ErrorCode;
use Mnb\PHPExcel\Support\MnbExcelException;

final class FormulaGuard
{
    /** @var list<string> */
    private const DANGEROUS_FUNCTIONS = [
        'CALL',
        'REGISTER.ID',
        'EXEC',
        'HYPERLINK',
        'WEBSERVICE',
        'FILTERXML',
        'DDE',
        'RTD',
        'IMPORTXML',
        'IMPORTDATA',
        'IMPORTRANGE',
        'IMPORTHTML',
    ];

    /**
     * Validate explicit formulas before XLSX writing.
     *
     * @param array<string,mixed> $options
     */
    public static function assertSafe(string $formula, array $options = []): void
    {
        $policy = strtolower((string) ($options['formula_policy'] ?? 'safe'));
        if ($policy === 'allow') {
            return;
        }
        if ($policy === 'block') {
            throw MnbExcelException::withCode('Formula writing is blocked by formula_policy=block.', ErrorCode::SECURITY_BLOCKED);
        }
        if ($policy !== 'safe') {
            throw MnbExcelException::withCode('Unknown formula policy: ' . $policy, ErrorCode::INVALID_ARGUMENT);
        }

        $risk = self::risk($formula, $options);
        if ($risk['risk'] !== 'none') {
            throw MnbExcelException::withCode('Unsafe formula blocked: ' . implode('; ', $risk['reasons']), ErrorCode::SECURITY_BLOCKED);
        }
    }

    /**
     * @param array<string,mixed> $options
     * @return array{risk:string,reasons:list<string>}
     */
    public static function risk(string $formula, array $options = []): array
    {
        $formula = trim($formula);
        if (str_starts_with($formula, '=')) {
            $formula = substr($formula, 1);
        }

        $reasons = [];
        if ($formula === '') {
            $reasons[] = 'empty formula';
        }
        if (preg_match('/[\x00-\x08\x0B\x0C\x0E-\x1F]/', $formula) === 1) {
            $reasons[] = 'control characters';
        }
        if (preg_match('/\[[^\]]+\]/', $formula) === 1) {
            $reasons[] = 'external workbook reference';
        }
        if (preg_match('/(?:https?|ftp|file):\/\//i', $formula) === 1) {
            $reasons[] = 'external URL reference';
        }
        if (preg_match('/(?:^|[^A-Z0-9_.])(?:' . implode('|', array_map('preg_quote', self::DANGEROUS_FUNCTIONS)) . ')\s*\(/i', $formula) === 1) {
            $reasons[] = 'dangerous function';
        }
        if (($options['allow_sheet_references'] ?? true) === false && preg_match('/(?:^|[^\'\"])(?:[A-Za-z0-9_ ]+!)\$?[A-Z]+\$?\d+/u', $formula) === 1) {
            $reasons[] = 'sheet reference';
        }

        return [
            'risk' => $reasons === [] ? 'none' : 'high',
            'reasons' => $reasons,
        ];
    }
}
