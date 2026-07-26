<?php

declare(strict_types=1);

namespace Mnb\PHPExcel\Security;

use Mnb\PHPExcel\Support\ErrorCode;
use Mnb\PHPExcel\Support\MnbExcelException;

/** ECMA-376 Standard Encryption (AES-128/SHA-1), retained for broad viewer compatibility. */
final class StandardXlsxEncryption
{
    private const SPIN_COUNT = 50000;
    private const PROVIDER = 'Microsoft Enhanced RSA and AES Cryptographic Provider';

    public function encrypt(string $package, string $password): string
    {
        $this->assertRuntime();
        $this->assertPassword($password);
        $salt = random_bytes(16);
        $verifier = random_bytes(16);
        $key = $this->key($password, $salt);
        $encryptedVerifier = $this->aes($verifier, $key, true, false);
        $verifierHash = hash('sha1', $verifier, true);
        $encryptedVerifierHash = $this->aes(str_pad($verifierHash, 32, "\0"), $key, true, false);
        $encryptedPackage = $this->aes($package, $key, true, true);

        $provider = $this->utf16le(self::PROVIDER . "\0");
        $header = pack('V', 0x24)
            . pack('V', 0)
            . pack('V', 0x0000660E)
            . pack('V', 0x00008004)
            . pack('V', 128)
            . pack('V', 0x18)
            . pack('V', 0)
            . pack('V', 0)
            . $provider;
        $verifierBlock = pack('V', 16)
            . $salt
            . $encryptedVerifier
            . pack('V', 20)
            . $encryptedVerifierHash;
        $info = pack('v', 3) . pack('v', 2) . pack('V', 0x24) . pack('V', strlen($header)) . $header . $verifierBlock;

        return CompoundFile::write([
            'EncryptedPackage' => $this->packU64(strlen($package)) . $encryptedPackage,
            'EncryptionInfo' => $info,
        ]);
    }

    public function decrypt(string $container, string $password, array $options = []): string
    {
        $this->assertRuntime();
        $this->assertPassword($password);
        $streams = CompoundFile::read($container);
        if (!isset($streams['EncryptionInfo'], $streams['EncryptedPackage'])) {
            throw new MnbExcelException('Encrypted Office package streams are missing.');
        }
        $info = $streams['EncryptionInfo'];
        if (strlen($info) < 84 || unpack('v', substr($info, 0, 2))[1] !== 3 || unpack('v', substr($info, 2, 2))[1] !== 2) {
            throw new MnbExcelException('File is not ECMA-376 Standard Encryption.');
        }
        $headerSize = $this->u32($info, 8);
        $headerOffset = 12;
        $algId = $this->u32($info, $headerOffset + 8);
        $hashId = $this->u32($info, $headerOffset + 12);
        $keyBits = $this->u32($info, $headerOffset + 16);
        if ($algId !== 0x0000660E || $hashId !== 0x00008004 || $keyBits !== 128) {
            throw new MnbExcelException('Unsupported Standard Encryption algorithm. AES-128 with SHA-1 is required.');
        }
        $verifierOffset = $headerOffset + $headerSize;
        $saltSize = $this->u32($info, $verifierOffset);
        if ($saltSize !== 16 || $verifierOffset + 40 > strlen($info)) {
            throw new MnbExcelException('Invalid Standard Encryption verifier.');
        }
        $salt = substr($info, $verifierOffset + 4, 16);
        $encryptedVerifier = substr($info, $verifierOffset + 20, 16);
        $hashSize = $this->u32($info, $verifierOffset + 36);
        $encryptedVerifierHash = substr($info, $verifierOffset + 40, 32);
        $key = $this->key($password, $salt);
        $verifier = $this->aes($encryptedVerifier, $key, false, false);
        $verifierHash = $this->aes($encryptedVerifierHash, $key, false, false);
        if (!hash_equals(substr($verifierHash, 0, $hashSize), hash('sha1', $verifier, true))) {
            throw MnbExcelException::withCode('Incorrect password for encrypted XLSX file.', ErrorCode::FILE_READ_FAILED, ['encrypted' => true]);
        }

        $encryptedPackage = $streams['EncryptedPackage'];
        if (strlen($encryptedPackage) < 8) {
            throw new MnbExcelException('EncryptedPackage stream is truncated.');
        }
        $plainSize = $this->u64($encryptedPackage, 0);
        $maxBytes = max(1, (int) ($options['max_decrypted_bytes'] ?? 2_147_483_648));
        if ($plainSize > $maxBytes) {
            throw new MnbExcelException('Decrypted workbook exceeds max_decrypted_bytes.');
        }
        $plain = $this->aes(substr($encryptedPackage, 8), $key, false, false);
        if (strlen($plain) < $plainSize) {
            throw new MnbExcelException('Decrypted XLSX package is truncated.');
        }
        return substr($plain, 0, $plainSize);
    }

    private function key(string $password, string $salt): string
    {
        $hash = hash('sha1', $salt . $this->utf16le($password), true);
        for ($i = 0; $i < self::SPIN_COUNT; $i++) {
            $hash = hash('sha1', pack('V', $i) . $hash, true);
        }
        $finalHash = hash('sha1', $hash . pack('V', 0), true);
        $x1 = str_repeat("\x36", 64);
        $x2 = str_repeat("\x5C", 64);
        for ($i = 0; $i < strlen($finalHash); $i++) {
            $x1[$i] = chr(ord($x1[$i]) ^ ord($finalHash[$i]));
            $x2[$i] = chr(ord($x2[$i]) ^ ord($finalHash[$i]));
        }
        return substr(hash('sha1', $x1, true) . hash('sha1', $x2, true), 0, 16);
    }

    private function aes(string $data, string $key, bool $encrypt, bool $padding): string
    {
        $flags = OPENSSL_RAW_DATA | ($padding ? 0 : OPENSSL_ZERO_PADDING);
        $result = $encrypt
            ? openssl_encrypt($data, 'aes-128-ecb', $key, $flags)
            : openssl_decrypt($data, 'aes-128-ecb', $key, $flags);
        if ($result === false) {
            throw new MnbExcelException('OpenSSL AES operation failed.');
        }
        return $result;
    }

    private function utf16le(string $value): string
    {
        $converted = function_exists('iconv') ? iconv('UTF-8', 'UTF-16LE//IGNORE', $value) : false;
        if ($converted === false) {
            throw new MnbExcelException('iconv is required for Office document encryption.');
        }
        return $converted;
    }

    private function assertRuntime(): void
    {
        if (!extension_loaded('openssl')) {
            throw MnbExcelException::withCode('ext-openssl is required for encrypted XLSX files.', ErrorCode::EXTENSION_MISSING, ['extension' => 'openssl']);
        }
    }

    private function assertPassword(string $password): void
    {
        if ($password === '') {
            throw new MnbExcelException('Encryption password cannot be empty.');
        }
        $count = preg_match_all('/./us', $password, $matches);
        if (($count === false ? strlen($password) : $count) > 255) {
            throw new MnbExcelException('Encryption password cannot exceed 255 Unicode code points.');
        }
    }

    private function u32(string $data, int $offset): int
    {
        return unpack('V', substr($data, $offset, 4))[1];
    }

    private function u64(string $data, int $offset): int
    {
        $parts = unpack('Vlow/Vhigh', substr($data, $offset, 8));
        return (int) ($parts['low'] + ($parts['high'] * 4294967296));
    }

    private function packU64(int $value): string
    {
        return pack('V2', $value & 0xFFFFFFFF, intdiv($value, 4294967296));
    }
}
