<?php

declare(strict_types=1);

namespace Mnb\PHPExcel\Security;

use Mnb\PHPExcel\Support\MnbExcelException;

/**
 * Minimal Compound Binary File (OLE/CFB) reader/writer used for encrypted OOXML packages.
 * Supports CFB v3/v4 reading and emits interoperable CFB v3 containers.
 */
final class CompoundFile
{
    private const MAGIC = "\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1";
    private const FREESECT = 0xFFFFFFFF;
    private const ENDOFCHAIN = 0xFFFFFFFE;
    private const FATSECT = 0xFFFFFFFD;
    private const DIFSECT = 0xFFFFFFFC;
    private const NOSTREAM = 0xFFFFFFFF;

    /** @param array<string,string> $streams */
    public static function write(array $streams): string
    {
        if ($streams === []) {
            throw new MnbExcelException('Compound file requires at least one stream.');
        }

        $sectorSize = 512;
        $miniSectorSize = 64;
        $miniCutoff = 4096;
        $entriesPerFat = intdiv($sectorSize, 4);

        ksort($streams, SORT_STRING);
        $small = [];
        $large = [];
        foreach ($streams as $name => $data) {
            if ($name === '' || strlen(self::utf16le($name . "\0")) > 64) {
                throw new MnbExcelException('Invalid CFB stream name: ' . $name);
            }
            if (strlen($data) < $miniCutoff) {
                $small[$name] = $data;
            } else {
                $large[$name] = $data;
            }
        }

        $miniStream = '';
        $miniStarts = [];
        $miniFat = [];
        $miniId = 0;
        foreach ($small as $name => $data) {
            $count = max(1, (int) ceil(strlen($data) / $miniSectorSize));
            $miniStarts[$name] = $miniId;
            for ($i = 0; $i < $count; $i++, $miniId++) {
                $miniFat[$miniId] = $i === $count - 1 ? self::ENDOFCHAIN : $miniId + 1;
                $miniStream .= str_pad(substr($data, $i * $miniSectorSize, $miniSectorSize), $miniSectorSize, "\0");
            }
        }

        $sectors = [];
        $chains = [];
        $appendNormal = static function (string $data) use (&$sectors, &$chains, $sectorSize): int {
            if ($data === '') {
                return CompoundFile::ENDOFCHAIN;
            }
            $start = count($sectors);
            $count = (int) ceil(strlen($data) / $sectorSize);
            $ids = [];
            for ($i = 0; $i < $count; $i++) {
                $ids[] = count($sectors);
                $sectors[] = str_pad(substr($data, $i * $sectorSize, $sectorSize), $sectorSize, "\0");
            }
            $chains[] = $ids;
            return $start;
        };

        $normalStarts = [];
        foreach ($large as $name => $data) {
            $normalStarts[$name] = $appendNormal($data);
        }
        $rootMiniStart = $appendNormal($miniStream);

        $miniFatStart = self::ENDOFCHAIN;
        $miniFatSectorIds = [];
        if ($miniFat !== []) {
            $miniFatBinary = '';
            foreach ($miniFat as $value) {
                $miniFatBinary .= pack('V', $value);
            }
            $miniFatBinary = str_pad($miniFatBinary, (int) ceil(strlen($miniFatBinary) / $sectorSize) * $sectorSize, "\xFF");
            $miniFatStart = count($sectors);
            $count = intdiv(strlen($miniFatBinary), $sectorSize);
            for ($i = 0; $i < $count; $i++) {
                $miniFatSectorIds[] = count($sectors);
                $sectors[] = substr($miniFatBinary, $i * $sectorSize, $sectorSize);
            }
            $chains[] = $miniFatSectorIds;
        }

        $directoryEntryCount = count($streams) + 1;
        $directorySectorCount = max(1, (int) ceil(($directoryEntryCount * 128) / $sectorSize));
        $directorySectorId = count($sectors);
        $directorySectorIds = [];
        for ($i = 0; $i < $directorySectorCount; $i++) {
            $directorySectorIds[] = count($sectors);
            $sectors[] = str_repeat("\0", $sectorSize);
        }
        $chains[] = $directorySectorIds;

        $baseSectorCount = count($sectors);
        $fatCount = 0;
        $difatCount = 0;
        $difatCapacity = $entriesPerFat - 1;
        do {
            $previousFat = $fatCount;
            $previousDifat = $difatCount;
            $fatCount = (int) ceil(($baseSectorCount + $fatCount + $difatCount) / $entriesPerFat);
            $difatCount = $fatCount > 109 ? (int) ceil(($fatCount - 109) / $difatCapacity) : 0;
        } while ($fatCount !== $previousFat || $difatCount !== $previousDifat);

        $difatSectorIds = [];
        for ($i = 0; $i < $difatCount; $i++) {
            $difatSectorIds[] = count($sectors);
            $sectors[] = str_repeat("\xFF", $sectorSize);
        }

        $fatSectorIds = [];
        for ($i = 0; $i < $fatCount; $i++) {
            $fatSectorIds[] = count($sectors);
            $sectors[] = str_repeat("\xFF", $sectorSize);
        }

        $fat = array_fill(0, count($sectors), self::FREESECT);
        foreach ($chains as $ids) {
            foreach ($ids as $i => $sid) {
                $fat[$sid] = $i === count($ids) - 1 ? self::ENDOFCHAIN : $ids[$i + 1];
            }
        }
        foreach ($difatSectorIds as $sid) {
            $fat[$sid] = self::DIFSECT;
        }
        foreach ($fatSectorIds as $sid) {
            $fat[$sid] = self::FATSECT;
        }

        $fatBinary = '';
        foreach ($fat as $entry) {
            $fatBinary .= pack('V', $entry);
        }
        $fatBinary = str_pad($fatBinary, $fatCount * $sectorSize, "\xFF");
        foreach ($fatSectorIds as $i => $sid) {
            $sectors[$sid] = substr($fatBinary, $i * $sectorSize, $sectorSize);
        }

        $remainingFatSectorIds = array_slice($fatSectorIds, 109);
        foreach ($difatSectorIds as $i => $sid) {
            $chunk = array_slice($remainingFatSectorIds, $i * $difatCapacity, $difatCapacity);
            $binary = '';
            for ($j = 0; $j < $difatCapacity; $j++) {
                $binary .= pack('V', $chunk[$j] ?? self::FREESECT);
            }
            $binary .= pack('V', $difatSectorIds[$i + 1] ?? self::ENDOFCHAIN);
            $sectors[$sid] = $binary;
        }

        $directoryEntries = [];
        $names = array_keys($streams);
        $rootChild = $names === [] ? self::NOSTREAM : 1;
        $directoryEntries[] = self::directoryEntry(
            'Root Entry',
            5,
            self::NOSTREAM,
            self::NOSTREAM,
            $rootChild,
            $rootMiniStart,
            strlen($miniStream)
        );
        foreach ($names as $index => $name) {
            $id = $index + 1;
            $right = $id < count($names) ? $id + 1 : self::NOSTREAM;
            $data = $streams[$name];
            $start = isset($miniStarts[$name]) ? $miniStarts[$name] : ($normalStarts[$name] ?? self::ENDOFCHAIN);
            $directoryEntries[] = self::directoryEntry($name, 2, self::NOSTREAM, $right, self::NOSTREAM, $start, strlen($data));
        }
        $directory = str_pad(implode('', $directoryEntries), $directorySectorCount * $sectorSize, "\0");
        foreach ($directorySectorIds as $i => $sid) {
            $sectors[$sid] = substr($directory, $i * $sectorSize, $sectorSize);
        }

        $header = self::MAGIC
            . str_repeat("\0", 16)
            . pack('v', 0x003E)
            . pack('v', 3)
            . pack('v', 0xFFFE)
            . pack('v', 9)
            . pack('v', 6)
            . str_repeat("\0", 6)
            . pack('V', 0)
            . pack('V', $fatCount)
            . pack('V', $directorySectorId)
            . pack('V', 0)
            . pack('V', $miniCutoff)
            . pack('V', $miniFatStart)
            . pack('V', count($miniFatSectorIds))
            . pack('V', $difatSectorIds[0] ?? self::ENDOFCHAIN)
            . pack('V', count($difatSectorIds));
        for ($i = 0; $i < 109; $i++) {
            $header .= pack('V', $fatSectorIds[$i] ?? self::FREESECT);
        }
        $header = str_pad($header, $sectorSize, "\0");

        return $header . implode('', $sectors);
    }

    /** @return array<string,string> */
    public static function read(string $binary): array
    {
        if (strlen($binary) < 512 || substr($binary, 0, 8) !== self::MAGIC) {
            throw new MnbExcelException('File is not a valid OLE compound document.');
        }

        $major = self::u16($binary, 26);
        $sectorShift = self::u16($binary, 30);
        $miniShift = self::u16($binary, 32);
        $sectorSize = 1 << $sectorShift;
        $miniSectorSize = 1 << $miniShift;
        if (!in_array($major, [3, 4], true) || !in_array($sectorSize, [512, 4096], true) || $miniSectorSize !== 64) {
            throw new MnbExcelException('Unsupported OLE compound document version.');
        }
        $headerSize = $sectorSize;
        if (strlen($binary) < $headerSize) {
            throw new MnbExcelException('Truncated OLE compound document header.');
        }

        $fatCount = self::u32($binary, 44);
        $firstDir = self::u32($binary, 48);
        $miniCutoff = self::u32($binary, 56);
        $firstMiniFat = self::u32($binary, 60);
        $miniFatCount = self::u32($binary, 64);
        $firstDifat = self::u32($binary, 68);
        $difatCount = self::u32($binary, 72);

        $readSector = static function (int $sid) use ($binary, $headerSize, $sectorSize): string {
            if ($sid >= 0xFFFFFFF0) {
                throw new MnbExcelException('Invalid OLE sector chain.');
            }
            $offset = $headerSize + ($sid * $sectorSize);
            if ($offset < 0 || $offset + $sectorSize > strlen($binary)) {
                throw new MnbExcelException('OLE sector points outside the file.');
            }
            return substr($binary, $offset, $sectorSize);
        };

        $fatSectorIds = [];
        for ($i = 0; $i < 109; $i++) {
            $sid = self::u32($binary, 76 + ($i * 4));
            if ($sid !== self::FREESECT) {
                $fatSectorIds[] = $sid;
            }
        }
        $difatSid = $firstDifat;
        $difatEntries = intdiv($sectorSize, 4) - 1;
        for ($d = 0; $d < $difatCount && $difatSid !== self::ENDOFCHAIN; $d++) {
            $sector = $readSector($difatSid);
            for ($i = 0; $i < $difatEntries; $i++) {
                $sid = self::u32($sector, $i * 4);
                if ($sid !== self::FREESECT) {
                    $fatSectorIds[] = $sid;
                }
            }
            $difatSid = self::u32($sector, $sectorSize - 4);
        }
        $fatSectorIds = array_slice($fatSectorIds, 0, $fatCount);
        if (count($fatSectorIds) !== $fatCount) {
            throw new MnbExcelException('Invalid OLE FAT sector list.');
        }

        $fat = [];
        foreach ($fatSectorIds as $sid) {
            $sector = $readSector($sid);
            for ($i = 0; $i < $sectorSize; $i += 4) {
                $fat[] = self::u32($sector, $i);
            }
        }

        $readChain = static function (int $start, int $maxBytes = PHP_INT_MAX) use (&$fat, $readSector, $sectorSize): string {
            if ($start === CompoundFile::ENDOFCHAIN || $start === CompoundFile::FREESECT) {
                return '';
            }
            $out = '';
            $seen = [];
            $sid = $start;
            $limit = max(1, count($fat) + 1);
            while ($sid !== CompoundFile::ENDOFCHAIN) {
                if ($sid >= count($fat) || isset($seen[$sid]) || count($seen) > $limit) {
                    throw new MnbExcelException('Corrupt or cyclic OLE sector chain.');
                }
                $seen[$sid] = true;
                $out .= $readSector($sid);
                if (strlen($out) >= $maxBytes) {
                    break;
                }
                $sid = $fat[$sid];
            }
            return $out;
        };

        $directory = $readChain($firstDir);
        $entries = [];
        for ($offset = 0; $offset + 128 <= strlen($directory); $offset += 128) {
            $entry = substr($directory, $offset, 128);
            $nameLength = self::u16($entry, 64);
            $type = ord($entry[66]);
            if ($type === 0 || $nameLength < 2 || $nameLength > 64) {
                continue;
            }
            $name = self::utf16leDecode(substr($entry, 0, $nameLength - 2));
            $entries[] = [
                'name' => $name,
                'type' => $type,
                'start' => self::u32($entry, 116),
                'size' => self::u64($entry, 120),
            ];
        }
        $root = null;
        foreach ($entries as $entry) {
            if ($entry['type'] === 5) {
                $root = $entry;
                break;
            }
        }
        if ($root === null) {
            throw new MnbExcelException('OLE root directory entry is missing.');
        }

        $miniFat = [];
        if ($miniFatCount > 0 && $firstMiniFat !== self::ENDOFCHAIN) {
            $miniFatBinary = substr($readChain($firstMiniFat, $miniFatCount * $sectorSize), 0, $miniFatCount * $sectorSize);
            for ($i = 0; $i + 4 <= strlen($miniFatBinary); $i += 4) {
                $miniFat[] = self::u32($miniFatBinary, $i);
            }
        }
        $miniStream = $root['size'] > 0 ? substr($readChain($root['start'], $root['size']), 0, $root['size']) : '';

        $streams = [];
        foreach ($entries as $entry) {
            if ($entry['type'] !== 2) {
                continue;
            }
            $size = $entry['size'];
            if ($size < $miniCutoff && $miniFat !== []) {
                $out = '';
                $sid = $entry['start'];
                $seen = [];
                while ($sid !== self::ENDOFCHAIN && strlen($out) < $size) {
                    if ($sid >= count($miniFat) || isset($seen[$sid])) {
                        throw new MnbExcelException('Corrupt OLE mini-sector chain.');
                    }
                    $seen[$sid] = true;
                    $offset = $sid * $miniSectorSize;
                    if ($offset + $miniSectorSize > strlen($miniStream)) {
                        throw new MnbExcelException('OLE mini-sector points outside the mini stream.');
                    }
                    $out .= substr($miniStream, $offset, $miniSectorSize);
                    $sid = $miniFat[$sid];
                }
                $streams[$entry['name']] = substr($out, 0, $size);
            } else {
                $streams[$entry['name']] = substr($readChain($entry['start'], $size), 0, $size);
            }
        }
        return $streams;
    }

    private static function directoryEntry(string $name, int $type, int $left, int $right, int $child, int $start, int $size): string
    {
        $encodedName = self::utf16le($name . "\0");
        $entry = str_pad($encodedName, 64, "\0")
            . pack('v', strlen($encodedName))
            . chr($type)
            . chr(1)
            . pack('V', $left)
            . pack('V', $right)
            . pack('V', $child)
            . str_repeat("\0", 16)
            . pack('V', 0)
            . str_repeat("\0", 16)
            . pack('V', $start)
            . self::packU64($size);
        return str_pad($entry, 128, "\0");
    }

    private static function utf16le(string $value): string
    {
        $converted = function_exists('iconv') ? iconv('UTF-8', 'UTF-16LE//IGNORE', $value) : false;
        if ($converted === false) {
            throw new MnbExcelException('iconv is required for Office document encryption.');
        }
        return $converted;
    }

    private static function utf16leDecode(string $value): string
    {
        $converted = function_exists('iconv') ? iconv('UTF-16LE', 'UTF-8//IGNORE', $value) : false;
        return $converted === false ? '' : $converted;
    }

    private static function u16(string $data, int $offset): int
    {
        return unpack('v', substr($data, $offset, 2))[1];
    }

    private static function u32(string $data, int $offset): int
    {
        return unpack('V', substr($data, $offset, 4))[1];
    }

    private static function u64(string $data, int $offset): int
    {
        $parts = unpack('Vlow/Vhigh', substr($data, $offset, 8));
        if ($parts['high'] > 0x7FFFFFFF || ($parts['high'] > 0 && PHP_INT_SIZE < 8)) {
            throw new MnbExcelException('OLE stream is too large for this PHP runtime.');
        }
        return (int) ($parts['low'] + ($parts['high'] * 4294967296));
    }

    private static function packU64(int $value): string
    {
        $low = $value & 0xFFFFFFFF;
        $high = intdiv($value, 4294967296);
        return pack('V2', $low, $high);
    }
}
