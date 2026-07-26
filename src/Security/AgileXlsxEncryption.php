<?php

declare(strict_types=1);

namespace Mnb\PHPExcel\Security;

use Mnb\PHPExcel\Support\ErrorCode;
use Mnb\PHPExcel\Support\MnbExcelException;

/** Standards-compatible ECMA-376 Agile Encryption for OOXML packages. */
final class AgileXlsxEncryption
{
    private const BLOCK_VERIFIER_INPUT = "\xFE\xA7\xD2\x76\x3B\x4B\x9E\x79";
    private const BLOCK_VERIFIER_HASH = "\xD7\xAA\x0F\x6D\x30\x61\x34\x4E";
    private const BLOCK_CRYPTO_KEY = "\x14\x6E\x0B\xE7\xAB\xAC\xD0\xD6";
    private const BLOCK_INTEGRITY_KEY = "\x5F\xB2\xAD\x01\x0C\xB9\xE1\xF6";
    private const BLOCK_INTEGRITY_VALUE = "\xA0\x67\x7F\x02\xB2\x2C\x84\x33";
    private const CHUNK_SIZE = 4096;

    /** @param array<string,mixed> $options */
    public function encrypt(string $package, string $password, array $options = []): string
    {
        $this->assertRuntime();
        $this->assertPassword($password);
        $spinCount = max(1000, min(10_000_000, (int) ($options['spin_count'] ?? 100000)));
        $keyBits = 256;
        $keyBytes = 32;
        $blockSize = 16;
        $hashSize = 64;
        $passwordSalt = random_bytes(16);
        $packageSalt = random_bytes(16);
        $packageKey = random_bytes($keyBytes);
        $verifier = random_bytes(16);

        $passwordHash = $this->passwordHash($password, $passwordSalt, $spinCount);
        $encryptedVerifier = $this->crypt(
            $verifier,
            $this->deriveKey($passwordHash, self::BLOCK_VERIFIER_INPUT, $keyBytes),
            $this->deriveIv($passwordSalt, null, $blockSize),
            true
        );
        $encryptedVerifierHash = $this->crypt(
            hash('sha512', $verifier, true),
            $this->deriveKey($passwordHash, self::BLOCK_VERIFIER_HASH, $keyBytes),
            $this->deriveIv($passwordSalt, null, $blockSize),
            true
        );
        $encryptedPackageKey = $this->crypt(
            $packageKey,
            $this->deriveKey($passwordHash, self::BLOCK_CRYPTO_KEY, $keyBytes),
            $this->deriveIv($passwordSalt, null, $blockSize),
            true
        );

        $encryptedPayload = '';
        $packageLength = strlen($package);
        for ($offset = 0, $block = 0; $offset < $packageLength; $offset += self::CHUNK_SIZE, $block++) {
            $chunk = substr($package, $offset, self::CHUNK_SIZE);
            $chunk = $this->padZero($chunk, $blockSize);
            $iv = $this->deriveIv($packageSalt, pack('V', $block), $blockSize);
            $encryptedPayload .= $this->crypt($chunk, $packageKey, $iv, true);
        }
        $encryptedPackage = $this->packU64($packageLength) . $encryptedPayload;

        $hmacKey = random_bytes($hashSize);
        $hmacValue = hash_hmac('sha512', $encryptedPackage, $hmacKey, true);
        $encryptedHmacKey = $this->crypt(
            $this->padZero($hmacKey, $blockSize),
            $packageKey,
            $this->deriveIv($packageSalt, self::BLOCK_INTEGRITY_KEY, $blockSize),
            true
        );
        $encryptedHmacValue = $this->crypt(
            $this->padZero($hmacValue, $blockSize),
            $packageKey,
            $this->deriveIv($packageSalt, self::BLOCK_INTEGRITY_VALUE, $blockSize),
            true
        );

        $xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            . '<encryption xmlns="http://schemas.microsoft.com/office/2006/encryption" xmlns:p="http://schemas.microsoft.com/office/2006/keyEncryptor/password">'
            . '<keyData saltSize="16" blockSize="16" keyBits="' . $keyBits . '" hashSize="' . $hashSize . '" cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" hashAlgorithm="SHA512" saltValue="' . base64_encode($packageSalt) . '"/>'
            . '<dataIntegrity encryptedHmacKey="' . base64_encode($encryptedHmacKey) . '" encryptedHmacValue="' . base64_encode($encryptedHmacValue) . '"/>'
            . '<keyEncryptors><keyEncryptor uri="http://schemas.microsoft.com/office/2006/keyEncryptor/password">'
            . '<p:encryptedKey spinCount="' . $spinCount . '" saltSize="16" blockSize="16" keyBits="' . $keyBits . '" hashSize="' . $hashSize . '" cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" hashAlgorithm="SHA512" saltValue="' . base64_encode($passwordSalt) . '" encryptedVerifierHashInput="' . base64_encode($encryptedVerifier) . '" encryptedVerifierHashValue="' . base64_encode($encryptedVerifierHash) . '" encryptedKeyValue="' . base64_encode($encryptedPackageKey) . '"/>'
            . '</keyEncryptor></keyEncryptors></encryption>';

        $encryptionInfo = "\x04\x00\x04\x00\x40\x00\x00\x00" . $xml;
        return CompoundFile::write([
            'EncryptedPackage' => $encryptedPackage,
            'EncryptionInfo' => $encryptionInfo,
        ]);
    }

    /** @param array<string,mixed> $options */
    public function decrypt(string $container, string $password, array $options = []): string
    {
        $this->assertRuntime();
        $this->assertPassword($password);
        $streams = CompoundFile::read($container);
        if (!isset($streams['EncryptionInfo'], $streams['EncryptedPackage'])) {
            throw new MnbExcelException('Encrypted Office package streams are missing.');
        }
        $info = $streams['EncryptionInfo'];
        if (strlen($info) < 9 || substr($info, 0, 4) !== "\x04\x00\x04\x00") {
            throw new MnbExcelException('Only ECMA-376 Agile encrypted XLSX files are supported.');
        }
        $xml = substr($info, 8);
        $keyData = $this->tagAttributes($xml, 'keyData');
        $integrity = $this->tagAttributes($xml, 'dataIntegrity');
        $encryptedKey = $this->tagAttributes($xml, 'encryptedKey');
        if ($keyData === [] || $encryptedKey === []) {
            throw new MnbExcelException('Invalid Agile EncryptionInfo XML.');
        }

        $this->assertAlgorithm($keyData, $encryptedKey);
        $spinCount = (int) ($encryptedKey['spinCount'] ?? 0);
        $maxSpin = max(100000, (int) ($options['max_spin_count'] ?? 10_000_000));
        if ($spinCount < 1 || $spinCount > $maxSpin) {
            throw new MnbExcelException('Encrypted workbook spin count is outside the allowed security limit.');
        }
        $keyBytes = intdiv((int) ($encryptedKey['keyBits'] ?? 256), 8);
        $blockSize = (int) ($encryptedKey['blockSize'] ?? 16);
        $hashSize = (int) ($encryptedKey['hashSize'] ?? 64);
        $passwordSalt = $this->b64($encryptedKey['saltValue'] ?? '');
        $packageSalt = $this->b64($keyData['saltValue'] ?? '');
        $passwordHash = $this->passwordHash($password, $passwordSalt, $spinCount);

        $verifier = $this->crypt(
            $this->b64($encryptedKey['encryptedVerifierHashInput'] ?? ''),
            $this->deriveKey($passwordHash, self::BLOCK_VERIFIER_INPUT, $keyBytes),
            $this->deriveIv($passwordSalt, null, $blockSize),
            false
        );
        $verifierHash = $this->crypt(
            $this->b64($encryptedKey['encryptedVerifierHashValue'] ?? ''),
            $this->deriveKey($passwordHash, self::BLOCK_VERIFIER_HASH, $keyBytes),
            $this->deriveIv($passwordSalt, null, $blockSize),
            false
        );
        $expected = hash('sha512', substr($verifier, 0, 16), true);
        if (!hash_equals(substr($verifierHash, 0, $hashSize), substr($expected, 0, $hashSize))) {
            throw MnbExcelException::withCode('Incorrect password for encrypted XLSX file.', ErrorCode::FILE_READ_FAILED, ['encrypted' => true]);
        }

        $packageKey = $this->crypt(
            $this->b64($encryptedKey['encryptedKeyValue'] ?? ''),
            $this->deriveKey($passwordHash, self::BLOCK_CRYPTO_KEY, $keyBytes),
            $this->deriveIv($passwordSalt, null, $blockSize),
            false
        );
        $packageKey = substr($packageKey, 0, $keyBytes);

        $encryptedPackage = $streams['EncryptedPackage'];
        if (strlen($encryptedPackage) < 8) {
            throw new MnbExcelException('EncryptedPackage stream is truncated.');
        }
        $plainSize = $this->u64($encryptedPackage, 0);
        $maxBytes = max(1, (int) ($options['max_decrypted_bytes'] ?? 2_147_483_648));
        if ($plainSize > $maxBytes) {
            throw new MnbExcelException('Decrypted workbook exceeds max_decrypted_bytes.');
        }

        if ($integrity !== []) {
            $hmacKey = $this->crypt(
                $this->b64($integrity['encryptedHmacKey'] ?? ''),
                $packageKey,
                $this->deriveIv($packageSalt, self::BLOCK_INTEGRITY_KEY, $blockSize),
                false
            );
            $hmacValue = $this->crypt(
                $this->b64($integrity['encryptedHmacValue'] ?? ''),
                $packageKey,
                $this->deriveIv($packageSalt, self::BLOCK_INTEGRITY_VALUE, $blockSize),
                false
            );
            $actual = hash_hmac('sha512', $encryptedPackage, substr($hmacKey, 0, $hashSize), true);
            if (!hash_equals(substr($hmacValue, 0, $hashSize), $actual)) {
                throw new MnbExcelException('Encrypted XLSX integrity verification failed.');
            }
        }

        $ciphertext = substr($encryptedPackage, 8);
        $plain = '';
        $remaining = $plainSize;
        $cipherOffset = 0;
        for ($block = 0; $remaining > 0; $block++) {
            $plainChunkLength = min(self::CHUNK_SIZE, $remaining);
            $cipherChunkLength = (int) ceil($plainChunkLength / $blockSize) * $blockSize;
            $chunk = substr($ciphertext, $cipherOffset, $cipherChunkLength);
            if (strlen($chunk) !== $cipherChunkLength) {
                throw new MnbExcelException('Encrypted XLSX package is truncated.');
            }
            $iv = $this->deriveIv($packageSalt, pack('V', $block), $blockSize);
            $plain .= substr($this->crypt($chunk, $packageKey, $iv, false), 0, $plainChunkLength);
            $cipherOffset += $cipherChunkLength;
            $remaining -= $plainChunkLength;
        }
        return $plain;
    }

    public function isEncrypted(string $binary): bool
    {
        return str_starts_with($binary, "\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1");
    }

    private function passwordHash(string $password, string $salt, int $spinCount): string
    {
        $passwordUtf16 = function_exists('iconv') ? iconv('UTF-8', 'UTF-16LE//IGNORE', $password) : false;
        if ($passwordUtf16 === false) {
            throw new MnbExcelException('iconv is required for Office document encryption.');
        }
        $hash = hash('sha512', $salt . $passwordUtf16, true);
        for ($i = 0; $i < $spinCount; $i++) {
            $hash = hash('sha512', pack('V', $i) . $hash, true);
        }
        return $hash;
    }

    private function deriveKey(string $passwordHash, string $blockKey, int $keyBytes): string
    {
        return substr(str_pad(hash('sha512', $passwordHash . $blockKey, true), $keyBytes, "\x36"), 0, $keyBytes);
    }

    private function deriveIv(string $salt, ?string $blockKey, int $blockSize): string
    {
        $value = $blockKey === null ? $salt : hash('sha512', $salt . $blockKey, true);
        return substr(str_pad($value, $blockSize, "\x36"), 0, $blockSize);
    }

    private function crypt(string $data, string $key, string $iv, bool $encrypt): string
    {
        $method = match (strlen($key)) {
            16 => 'aes-128-cbc',
            24 => 'aes-192-cbc',
            32 => 'aes-256-cbc',
            default => throw new MnbExcelException('Unsupported Agile encryption key size.'),
        };
        $result = $encrypt
            ? openssl_encrypt($data, $method, $key, OPENSSL_RAW_DATA | OPENSSL_ZERO_PADDING, $iv)
            : openssl_decrypt($data, $method, $key, OPENSSL_RAW_DATA | OPENSSL_ZERO_PADDING, $iv);
        if ($result === false) {
            throw new MnbExcelException('OpenSSL AES operation failed.');
        }
        return $result;
    }

    private function padZero(string $data, int $blockSize): string
    {
        $remainder = strlen($data) % $blockSize;
        return $remainder === 0 ? $data : $data . str_repeat("\0", $blockSize - $remainder);
    }

    /** @return array<string,string> */
    private function tagAttributes(string $xml, string $localName): array
    {
        if (preg_match('/<(?:[A-Za-z_][A-Za-z0-9_.-]*:)?' . preg_quote($localName, '/') . '\b([^>]*)\/?\s*>/is', $xml, $match) !== 1) {
            return [];
        }
        $attrs = [];
        preg_match_all('/([A-Za-z0-9_:\-]+)\s*=\s*("([^"]*)"|\'([^\']*)\')/u', $match[1], $matches, PREG_SET_ORDER);
        foreach ($matches as $item) {
            $name = str_contains($item[1], ':') ? substr($item[1], strrpos($item[1], ':') + 1) : $item[1];
            $attrs[$name] = html_entity_decode($item[3] !== '' ? $item[3] : $item[4], ENT_QUOTES | ENT_XML1, 'UTF-8');
        }
        return $attrs;
    }

    /** @param array<string,string> $keyData @param array<string,string> $encryptedKey */
    private function assertAlgorithm(array $keyData, array $encryptedKey): void
    {
        foreach ([$keyData, $encryptedKey] as $attrs) {
            if (strtoupper($attrs['cipherAlgorithm'] ?? '') !== 'AES'
                || strtoupper($attrs['hashAlgorithm'] ?? '') !== 'SHA512'
                || strcasecmp($attrs['cipherChaining'] ?? '', 'ChainingModeCBC') !== 0
                || !in_array((int) ($attrs['keyBits'] ?? 0), [128, 192, 256], true)
            ) {
                throw new MnbExcelException('Unsupported Agile encryption algorithm. AES-CBC with SHA-512 is required.');
            }
        }
    }

    private function b64(string $value): string
    {
        $decoded = base64_decode($value, true);
        if ($decoded === false) {
            throw new MnbExcelException('Invalid base64 value in EncryptionInfo.');
        }
        return $decoded;
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

    private function packU64(int $value): string
    {
        return pack('V2', $value & 0xFFFFFFFF, intdiv($value, 4294967296));
    }

    private function u64(string $data, int $offset): int
    {
        $parts = unpack('Vlow/Vhigh', substr($data, $offset, 8));
        return (int) ($parts['low'] + ($parts['high'] * 4294967296));
    }
}
