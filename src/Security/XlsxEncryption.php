<?php

declare(strict_types=1);

namespace Mnb\PHPExcel\Security;

use Mnb\PHPExcel\Support\AtomicFileWriter;
use Mnb\PHPExcel\Support\ErrorCode;
use Mnb\PHPExcel\Support\MnbExcelException;

final class XlsxEncryption
{
    public function __construct(
        private readonly AgileXlsxEncryption $agile = new AgileXlsxEncryption(),
        private readonly StandardXlsxEncryption $standard = new StandardXlsxEncryption()
    ) {
    }

    public function isEncryptedFile(string $path): bool
    {
        $handle = @fopen($path, 'rb');
        if ($handle === false) {
            return false;
        }
        $magic = fread($handle, 8);
        fclose($handle);
        return $magic === "\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1";
    }


    /** Return agile, standard, unknown, or null for an unencrypted file. */
    public function encryptionMode(string $path): ?string
    {
        if (!$this->isEncryptedFile($path)) {
            return null;
        }
        try {
            $streams = CompoundFile::read($this->readFile($path, 2_147_483_648));
            $info = $streams['EncryptionInfo'] ?? '';
            if (strlen($info) >= 4) {
                $major = unpack('v', substr($info, 0, 2))[1];
                $minor = unpack('v', substr($info, 2, 2))[1];
                if ($major === 3 && $minor === 2) {
                    return 'standard';
                }
                if ($major === 4 && $minor === 4) {
                    return 'agile';
                }
            }
        } catch (\Throwable) {
        }
        return 'unknown';
    }

    /** @param array<string,mixed> $options */
    public function encryptFile(string $source, string $destination, string $password, array $options = []): void
    {
        $data = $this->readFile($source, (int) ($options['max_source_bytes'] ?? 2_147_483_648));
        $mode = (bool) ($options['compatibility_mode'] ?? false)
            ? 'standard'
            : strtolower((string) ($options['mode'] ?? $options['encryption_mode'] ?? 'agile'));
        if (!in_array($mode, ['agile', 'standard'], true)) {
            throw new MnbExcelException('encryption mode must be agile or standard.');
        }
        $encrypted = $mode === 'standard'
            ? $this->standard->encrypt($data, $password)
            : $this->agile->encrypt($data, $password, $options);
        AtomicFileWriter::writeViaTemp($destination, static function (string $tmp) use ($encrypted): void {
            if (file_put_contents($tmp, $encrypted, LOCK_EX) === false) {
                throw new MnbExcelException('Unable to write encrypted XLSX file.');
            }
        });
    }

    /** @param array<string,mixed> $options */
    public function decryptFile(string $source, string $destination, string $password, array $options = []): void
    {
        $data = $this->readFile($source, (int) ($options['max_source_bytes'] ?? 2_147_483_648));
        $streams = CompoundFile::read($data);
        $info = $streams['EncryptionInfo'] ?? '';
        $major = strlen($info) >= 2 ? unpack('v', substr($info, 0, 2))[1] : 0;
        $minor = strlen($info) >= 4 ? unpack('v', substr($info, 2, 2))[1] : 0;
        $plain = ($major === 3 && $minor === 2)
            ? $this->standard->decrypt($data, $password, $options)
            : $this->agile->decrypt($data, $password, $options);
        AtomicFileWriter::writeViaTemp($destination, static function (string $tmp) use ($plain): void {
            if (file_put_contents($tmp, $plain, LOCK_EX) === false) {
                throw new MnbExcelException('Unable to write decrypted XLSX file.');
            }
        });
    }

    /** @param array<string,mixed> $options */
    public function decryptToTemporary(string $source, string $password, array $options = []): string
    {
        $tmp = tempnam((string) ($options['temp_directory'] ?? sys_get_temp_dir()), 'mnb_xlsx_dec_');
        if ($tmp === false) {
            throw new MnbExcelException('Unable to allocate temporary decrypted XLSX path.');
        }
        try {
            $this->decryptFile($source, $tmp, $password, $options);
            return $tmp;
        } catch (\Throwable $e) {
            @unlink($tmp);
            throw $e;
        }
    }

    private function readFile(string $path, int $maxBytes): string
    {
        $real = realpath($path);
        if ($real === false || !is_file($real)) {
            throw MnbExcelException::withCode('File not found: ' . $path, ErrorCode::FILE_NOT_FOUND, ['path' => $path]);
        }
        $size = filesize($real);
        if ($size !== false && $size > $maxBytes) {
            throw new MnbExcelException('Encrypted document exceeds the configured size limit.');
        }
        $data = file_get_contents($real);
        if ($data === false) {
            throw MnbExcelException::withCode('Unable to read file: ' . $path, ErrorCode::FILE_READ_FAILED, ['path' => $path]);
        }
        return $data;
    }
}
