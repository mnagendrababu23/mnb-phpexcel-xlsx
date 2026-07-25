<?php

declare(strict_types=1);

namespace Mnb\PHPExcel\Reader;

final class XlsxStyleMap
{
    /** @var array<int, string> */
    private array $customNumberFormats = [];

    /** @var array<int, int> style index => numFmtId */
    private array $cellStyleNumberFormatIds = [];

    public static function fromXml(?string $stylesXml): self
    {
        $map = new self();
        if ($stylesXml === null || trim($stylesXml) === '') {
            return $map;
        }

        $map->customNumberFormats = $map->parseCustomNumberFormats($stylesXml);
        $map->cellStyleNumberFormatIds = $map->parseCellXfs($stylesXml);

        return $map;
    }

    public function isDateStyle(?int $styleIndex): bool
    {
        if ($styleIndex === null) {
            return false;
        }

        $numFmtId = $this->cellStyleNumberFormatIds[$styleIndex] ?? null;
        if ($numFmtId === null) {
            return false;
        }

        if (self::isBuiltinDateNumberFormat($numFmtId)) {
            return true;
        }

        $formatCode = $this->customNumberFormats[$numFmtId] ?? null;
        if ($formatCode === null) {
            return false;
        }

        return self::looksLikeDateFormat($formatCode);
    }

    public function numberFormatCode(?int $styleIndex): ?string
    {
        if ($styleIndex === null) {
            return null;
        }

        $numFmtId = $this->cellStyleNumberFormatIds[$styleIndex] ?? null;
        if ($numFmtId === null) {
            return null;
        }

        return $this->customNumberFormats[$numFmtId] ?? self::builtinNumberFormatCode($numFmtId);
    }

    /** @return array<int, string> */
    private function parseCustomNumberFormats(string $xml): array
    {
        $formats = [];
        preg_match_all('/<numFmt\b[^>]*>/i', $xml, $matches);
        foreach ($matches[0] ?? [] as $tag) {
            $attrs = $this->parseAttributes($tag);
            if (!isset($attrs['numFmtId'], $attrs['formatCode'])) {
                continue;
            }
            $formats[(int) $attrs['numFmtId']] = html_entity_decode($attrs['formatCode'], ENT_QUOTES | ENT_XML1, 'UTF-8');
        }

        return $formats;
    }

    /** @return array<int, int> */
    private function parseCellXfs(string $xml): array
    {
        if (!preg_match('/<cellXfs\b[^>]*>(.*?)<\/cellXfs>/is', $xml, $match)) {
            return [];
        }

        $xfs = [];
        preg_match_all('/<xf\b[^>]*>/i', $match[1], $xfMatches);
        foreach ($xfMatches[0] ?? [] as $tag) {
            $attrs = $this->parseAttributes($tag);
            $xfs[] = isset($attrs['numFmtId']) ? (int) $attrs['numFmtId'] : 0;
        }

        return $xfs;
    }

    /** @return array<string, string> */
    private function parseAttributes(string $tag): array
    {
        $attrs = [];
        preg_match_all('/([A-Za-z0-9_:\-]+)\s*=\s*("([^"]*)"|\'([^\']*)\')/u', $tag, $matches, PREG_SET_ORDER);
        foreach ($matches as $match) {
            $attrs[$match[1]] = $match[3] !== '' ? $match[3] : $match[4];
        }

        return $attrs;
    }

    public static function isBuiltinDateNumberFormat(int $numFmtId): bool
    {
        return in_array($numFmtId, [14, 15, 16, 17, 18, 19, 20, 21, 22, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 45, 46, 47, 50, 51, 52, 53, 54, 55, 56, 57, 58], true);
    }

    public static function builtinNumberFormatCode(int $numFmtId): ?string
    {
        return match ($numFmtId) {
            14 => 'm/d/yy',
            15 => 'd-mmm-yy',
            16 => 'd-mmm',
            17 => 'mmm-yy',
            18 => 'h:mm AM/PM',
            19 => 'h:mm:ss AM/PM',
            20 => 'h:mm',
            21 => 'h:mm:ss',
            22 => 'm/d/yy h:mm',
            45 => 'mm:ss',
            46 => '[h]:mm:ss',
            47 => 'mmss.0',
            default => null,
        };
    }

    public static function looksLikeDateFormat(string $formatCode): bool
    {
        $format = strtolower($formatCode);
        $format = preg_replace('/"[^"]*"/', '', $format) ?? $format;
        $format = preg_replace('/\[[^\]]*\]/', '', $format) ?? $format;
        $format = str_replace(['\\', '_', '*'], '', $format);

        return preg_match('/(^|[^a-z])([ymdhse]|am\/pm)([^a-z]|$)/', $format) === 1
            && preg_match('/[ymdhs]/', $format) === 1;
    }
}
