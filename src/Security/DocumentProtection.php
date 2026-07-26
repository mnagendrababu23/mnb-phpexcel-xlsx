<?php

declare(strict_types=1);

namespace Mnb\PHPExcel\Security;

use Mnb\PHPExcel\Support\MnbExcelException;

/** Generates OOXML workbook/worksheet protection credentials and attributes. */
final class DocumentProtection
{
    /** @param array<string,mixed> $options @return array<string,string> */
    public static function sheetAttributes(string $password, array $options = []): array
    {
        $credentials = self::credentials($password, $options);
        $attrs = $credentials + ['sheet' => '1'];
        $defaults = [
            'objects' => true,
            'scenarios' => true,
            'formatCells' => true,
            'formatColumns' => true,
            'formatRows' => true,
            'insertColumns' => true,
            'insertRows' => true,
            'insertHyperlinks' => true,
            'deleteColumns' => true,
            'deleteRows' => true,
            'selectLockedCells' => false,
            'sort' => true,
            'autoFilter' => true,
            'pivotTables' => true,
            'selectUnlockedCells' => false,
        ];
        $permissions = array_replace($defaults, (array) ($options['restrictions'] ?? $options['permissions'] ?? []));
        foreach ($permissions as $name => $restricted) {
            $attrs[(string) $name] = $restricted ? '1' : '0';
        }
        return $attrs;
    }

    /** @param array<string,mixed> $options @return array<string,string> */
    public static function workbookAttributes(string $password, array $options = []): array
    {
        $credentials = self::credentials($password, $options, 'workbook');
        return $credentials + [
            'lockStructure' => (bool) ($options['lock_structure'] ?? true) ? '1' : '0',
            'lockWindows' => (bool) ($options['lock_windows'] ?? false) ? '1' : '0',
            'lockRevision' => (bool) ($options['lock_revision'] ?? false) ? '1' : '0',
        ];
    }

    /** @param array<string,mixed> $options @return array<string,string> */
    private static function credentials(string $password, array $options, string $prefix = ''): array
    {
        if ($password === '') {
            throw new MnbExcelException('Protection password cannot be empty.');
        }
        $algorithm = strtoupper((string) ($options['algorithm'] ?? 'SHA-512'));
        $phpAlgorithm = strtolower(str_replace('-', '', $algorithm));
        if (!in_array($phpAlgorithm, hash_algos(), true)) {
            throw new MnbExcelException('Unsupported document-protection hash algorithm: ' . $algorithm);
        }
        $spinCount = max(1, min(10_000_000, (int) ($options['spin_count'] ?? 100000)));
        $salt = isset($options['salt']) ? (string) $options['salt'] : random_bytes(16);
        if ($salt === '') {
            throw new MnbExcelException('Protection salt cannot be empty.');
        }
        $utf16 = function_exists('iconv') ? iconv('UTF-8', 'UTF-16LE//IGNORE', $password) : false;
        if ($utf16 === false) {
            throw new MnbExcelException('iconv is required for Excel document protection.');
        }
        $hash = hash($phpAlgorithm, $salt . $utf16, true);
        for ($i = 0; $i < $spinCount; $i++) {
            $hash = hash($phpAlgorithm, $hash . pack('V', $i), true);
        }

        $base = [
            'algorithmName' => $algorithm,
            'hashValue' => base64_encode($hash),
            'saltValue' => base64_encode($salt),
            'spinCount' => (string) $spinCount,
            'password' => self::legacyHash($password),
        ];
        if ($prefix === '') {
            return $base;
        }
        return [
            $prefix . 'AlgorithmName' => $base['algorithmName'],
            $prefix . 'HashValue' => $base['hashValue'],
            $prefix . 'SaltValue' => $base['saltValue'],
            $prefix . 'SpinCount' => $base['spinCount'],
            $prefix . 'Password' => $base['password'],
        ];
    }

    /** Legacy XOR hash retained for older spreadsheet applications. */
    public static function legacyHash(string $password): string
    {
        $bytes = function_exists('iconv') ? iconv('UTF-8', 'UTF-16LE//IGNORE', (function_exists('mb_substr') ? mb_substr($password, 0, 15, 'UTF-8') : substr($password, 0, 15))) : false;
        if ($bytes === false) {
            $bytes = substr($password, 0, 15);
        }
        // Excel's legacy verifier works on one-byte characters. Fold UTF-16LE low bytes.
        $chars = [];
        for ($i = 0; $i < strlen($bytes); $i += 2) {
            $chars[] = ord($bytes[$i]);
        }
        if ($chars === []) {
            $chars = array_map('ord', str_split(substr($password, 0, 15)));
        }
        $hash = 0;
        $length = count($chars);
        for ($i = $length - 1; $i >= 0; $i--) {
            $value = $chars[$i];
            $shift = $i + 1;
            $rotated = (($value << $shift) | ($value >> (15 - $shift))) & 0x7FFF;
            $hash ^= $rotated;
        }
        $hash ^= $length;
        $hash ^= 0xCE4B;
        return strtoupper(str_pad(dechex($hash & 0xFFFF), 4, '0', STR_PAD_LEFT));
    }
}
