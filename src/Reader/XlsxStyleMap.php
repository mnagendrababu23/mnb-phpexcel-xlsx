<?php

declare(strict_types=1);

namespace Mnb\PHPExcel\Reader;

final class XlsxStyleMap
{
    /** @var array<int,string> */
    private array $customNumberFormats = [];

    /** @var array<int,int> */
    private array $cellStyleNumberFormatIds = [];

    /** @var array<int,array<string,mixed>> */
    private array $styles = [];

    /** @var array<int,array<string,mixed>> */
    private array $fonts = [];

    /** @var array<int,array<string,mixed>> */
    private array $fills = [];

    /** @var array<int,array<string,mixed>> */
    private array $borders = [];

    public static function fromXml(?string $stylesXml): self
    {
        $map = new self();
        if ($stylesXml === null || trim($stylesXml) === '') {
            return $map;
        }

        $map->customNumberFormats = $map->parseCustomNumberFormats($stylesXml);
        $map->fonts = $map->parseFonts($stylesXml);
        $map->fills = $map->parseFills($stylesXml);
        $map->borders = $map->parseBorders($stylesXml);
        [$map->cellStyleNumberFormatIds, $map->styles] = $map->parseCellXfs($stylesXml);

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
        return $formatCode !== null && self::looksLikeDateFormat($formatCode);
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

    /** @return array<string,mixed> */
    public function styleForIndex(?int $styleIndex): array
    {
        if ($styleIndex === null) {
            return [];
        }
        return $this->styles[$styleIndex] ?? [];
    }

    /** @return array<int,array<string,mixed>> */
    public function allStyles(): array
    {
        return $this->styles;
    }

    /** @return array<int,string> */
    private function parseCustomNumberFormats(string $xml): array
    {
        $formats = [];
        preg_match_all('/<(?:[A-Za-z_][A-Za-z0-9_.-]*:)?numFmt\b[^>]*>/i', $xml, $matches);
        foreach ($matches[0] ?? [] as $tag) {
            $attrs = $this->parseAttributes($tag);
            if (!isset($attrs['numFmtId'], $attrs['formatCode'])) {
                continue;
            }
            $formats[(int) $attrs['numFmtId']] = html_entity_decode($attrs['formatCode'], ENT_QUOTES | ENT_XML1, 'UTF-8');
        }
        return $formats;
    }

    /** @return array<int,array<string,mixed>> */
    private function parseFonts(string $xml): array
    {
        $container = $this->container($xml, 'fonts');
        if ($container === null) {
            return [];
        }
        preg_match_all('/<(?:[A-Za-z_][A-Za-z0-9_.-]*:)?font\b[^>]*>.*?<\/(?:[A-Za-z_][A-Za-z0-9_.-]*:)?font>/is', $container, $matches);
        $fonts = [];
        foreach ($matches[0] ?? [] as $fontXml) {
            $font = [];
            foreach (['b' => 'bold', 'i' => 'italic', 'strike' => 'strike', 'outline' => 'outline', 'shadow' => 'shadow', 'condense' => 'condense', 'extend' => 'extend'] as $tag => $key) {
                if ($this->booleanTag($fontXml, $tag)) {
                    $font[$key] = true;
                }
            }
            foreach (['name' => 'name', 'rFont' => 'name', 'sz' => 'size', 'family' => 'family', 'charset' => 'charset', 'scheme' => 'scheme', 'vertAlign' => 'vertical_align', 'u' => 'underline'] as $tag => $key) {
                $attrs = $this->firstTagAttributes($fontXml, $tag);
                if ($attrs !== []) {
                    $value = $attrs['val'] ?? true;
                    $font[$key] = in_array($key, ['size'], true) && is_numeric($value) ? (float) $value : $value;
                }
            }
            $color = $this->firstTagAttributes($fontXml, 'color');
            if ($color !== []) {
                $font['color'] = $this->normalizeColor($color);
            }
            $fonts[] = $font;
        }
        return $fonts;
    }

    /** @return array<int,array<string,mixed>> */
    private function parseFills(string $xml): array
    {
        $container = $this->container($xml, 'fills');
        if ($container === null) {
            return [];
        }
        preg_match_all('/<(?:[A-Za-z_][A-Za-z0-9_.-]*:)?fill\b[^>]*>.*?<\/(?:[A-Za-z_][A-Za-z0-9_.-]*:)?fill>/is', $container, $matches);
        $fills = [];
        foreach ($matches[0] ?? [] as $fillXml) {
            $fill = [];
            if (preg_match('/<(?:[A-Za-z_][A-Za-z0-9_.-]*:)?patternFill\b([^>]*)>(.*?)<\/(?:[A-Za-z_][A-Za-z0-9_.-]*:)?patternFill>|<(?:[A-Za-z_][A-Za-z0-9_.-]*:)?patternFill\b([^>]*)\/>/is', $fillXml, $match) === 1) {
                $attrs = $this->parseAttributes('<patternFill ' . ($match[1] !== '' ? $match[1] : ($match[3] ?? '')) . '>');
                $fill['type'] = 'pattern';
                $fill['pattern'] = $attrs['patternType'] ?? 'none';
                $body = (string) ($match[2] ?? '');
                $fg = $this->firstTagAttributes($body, 'fgColor');
                $bg = $this->firstTagAttributes($body, 'bgColor');
                if ($fg !== []) {
                    $fill['foreground'] = $this->normalizeColor($fg);
                }
                if ($bg !== []) {
                    $fill['background'] = $this->normalizeColor($bg);
                }
            } elseif (preg_match('/<(?:[A-Za-z_][A-Za-z0-9_.-]*:)?gradientFill\b([^>]*)>(.*?)<\/(?:[A-Za-z_][A-Za-z0-9_.-]*:)?gradientFill>/is', $fillXml, $match) === 1) {
                $fill['type'] = 'gradient';
                $fill += $this->parseAttributes('<gradientFill ' . $match[1] . '>');
                preg_match_all('/<(?:[A-Za-z_][A-Za-z0-9_.-]*:)?stop\b([^>]*)>(.*?)<\/(?:[A-Za-z_][A-Za-z0-9_.-]*:)?stop>/is', $match[2], $stops, PREG_SET_ORDER);
                foreach ($stops as $stop) {
                    $stopAttrs = $this->parseAttributes('<stop ' . $stop[1] . '>');
                    $color = $this->firstTagAttributes($stop[2], 'color');
                    $fill['stops'][] = ['position' => isset($stopAttrs['position']) ? (float) $stopAttrs['position'] : null, 'color' => $this->normalizeColor($color)];
                }
            }
            $fills[] = $fill;
        }
        return $fills;
    }

    /** @return array<int,array<string,mixed>> */
    private function parseBorders(string $xml): array
    {
        $container = $this->container($xml, 'borders');
        if ($container === null) {
            return [];
        }
        preg_match_all('/<(?:[A-Za-z_][A-Za-z0-9_.-]*:)?border\b([^>]*)>.*?<\/(?:[A-Za-z_][A-Za-z0-9_.-]*:)?border>/is', $container, $matches, PREG_SET_ORDER);
        $borders = [];
        foreach ($matches as $borderMatch) {
            $borderXml = $borderMatch[0];
            $border = [];
            foreach ($this->parseAttributes('<border ' . $borderMatch[1] . '>') as $key => $value) {
                $border[$this->snake($key)] = $this->scalar($value);
            }
            foreach (['left', 'right', 'top', 'bottom', 'diagonal', 'vertical', 'horizontal', 'start', 'end'] as $side) {
                if (preg_match('/<(?:[A-Za-z_][A-Za-z0-9_.-]*:)?' . $side . '\b([^>]*)(?:\/>|>(.*?)<\/(?:[A-Za-z_][A-Za-z0-9_.-]*:)?' . $side . '>)/is', $borderXml, $sideMatch) !== 1) {
                    continue;
                }
                $sideAttrs = $this->parseAttributes('<' . $side . ' ' . $sideMatch[1] . '>');
                $sideData = [];
                if (isset($sideAttrs['style'])) {
                    $sideData['style'] = $sideAttrs['style'];
                }
                $color = $this->firstTagAttributes((string) ($sideMatch[2] ?? ''), 'color');
                if ($color !== []) {
                    $sideData['color'] = $this->normalizeColor($color);
                }
                if ($sideData !== []) {
                    $border[$side] = $sideData;
                }
            }
            $borders[] = $border;
        }
        return $borders;
    }

    /** @return array{array<int,int>,array<int,array<string,mixed>>} */
    private function parseCellXfs(string $xml): array
    {
        $container = $this->container($xml, 'cellXfs');
        if ($container === null) {
            return [[], []];
        }

        preg_match_all('/<(?:[A-Za-z_][A-Za-z0-9_.-]*:)?xf\b([^>]*?)(?:\/\s*>|>(.*?)<\/(?:[A-Za-z_][A-Za-z0-9_.-]*:)?xf\s*>)/is', $container, $matches, PREG_SET_ORDER);
        $numberFormats = [];
        $styles = [];
        foreach ($matches as $match) {
            $attrs = $this->parseAttributes('<xf ' . $match[1] . '>');
            $body = (string) ($match[2] ?? '');
            $numFmtId = isset($attrs['numFmtId']) ? (int) $attrs['numFmtId'] : 0;
            $numberFormats[] = $numFmtId;
            $style = [
                'style_index' => count($styles),
                'number_format_id' => $numFmtId,
                'number_format' => $this->customNumberFormats[$numFmtId] ?? self::builtinNumberFormatCode($numFmtId),
                'font' => $this->fonts[(int) ($attrs['fontId'] ?? 0)] ?? [],
                'fill' => $this->fills[(int) ($attrs['fillId'] ?? 0)] ?? [],
                'border' => $this->borders[(int) ($attrs['borderId'] ?? 0)] ?? [],
            ];
            foreach (['xfId', 'quotePrefix', 'pivotButton', 'applyNumberFormat', 'applyFont', 'applyFill', 'applyBorder', 'applyAlignment', 'applyProtection'] as $key) {
                if (isset($attrs[$key])) {
                    $style[$this->snake($key)] = $this->scalar($attrs[$key]);
                }
            }
            $alignment = $this->firstTagAttributes($body, 'alignment');
            if ($alignment !== []) {
                $style['alignment'] = array_combine(
                    array_map(fn (string $key): string => $this->snake($key), array_keys($alignment)),
                    array_map(fn (string $value): mixed => $this->scalar($value), array_values($alignment))
                ) ?: [];
            }
            $protection = $this->firstTagAttributes($body, 'protection');
            if ($protection !== []) {
                $style['protection'] = array_combine(
                    array_map(fn (string $key): string => $this->snake($key), array_keys($protection)),
                    array_map(fn (string $value): mixed => $this->scalar($value), array_values($protection))
                ) ?: [];
            }
            $styles[] = array_filter($style, static fn (mixed $value): bool => $value !== [] && $value !== null);
        }

        return [$numberFormats, $styles];
    }

    private function container(string $xml, string $tag): ?string
    {
        return preg_match('/<(?:[A-Za-z_][A-Za-z0-9_.-]*:)?' . preg_quote($tag, '/') . '\b[^>]*>(.*?)<\/(?:[A-Za-z_][A-Za-z0-9_.-]*:)?' . preg_quote($tag, '/') . '>/is', $xml, $match) === 1 ? $match[1] : null;
    }

    /** @return array<string,string> */
    private function firstTagAttributes(string $xml, string $tag): array
    {
        return preg_match('/<(?:[A-Za-z_][A-Za-z0-9_.-]*:)?' . preg_quote($tag, '/') . '\b([^>]*)\/?\s*>/is', $xml, $match) === 1
            ? $this->parseAttributes('<' . $tag . ' ' . $match[1] . '>')
            : [];
    }

    private function booleanTag(string $xml, string $tag): bool
    {
        if (preg_match('/<(?:[A-Za-z_][A-Za-z0-9_.-]*:)?' . preg_quote($tag, '/') . '\b([^>]*)\/?\s*>/is', $xml, $match) !== 1) {
            return false;
        }
        $attrs = $this->parseAttributes('<' . $tag . ' ' . $match[1] . '>');
        return !isset($attrs['val']) || !in_array(strtolower($attrs['val']), ['0', 'false', 'off', 'no'], true);
    }

    /** @return array<string,mixed> */
    private function normalizeColor(array $attrs): array
    {
        $color = [];
        foreach (['rgb', 'indexed', 'theme', 'tint', 'auto'] as $key) {
            if (isset($attrs[$key])) {
                $color[$key] = $this->scalar($attrs[$key]);
            }
        }
        return $color;
    }

    private function scalar(string $value): mixed
    {
        $lower = strtolower($value);
        if (in_array($lower, ['true', 'false'], true)) {
            return $lower === 'true';
        }
        if ($value === '1' || $value === '0') {
            return $value === '1';
        }
        if (is_numeric($value)) {
            return str_contains($value, '.') ? (float) $value : (int) $value;
        }
        return html_entity_decode($value, ENT_QUOTES | ENT_XML1, 'UTF-8');
    }

    private function snake(string $key): string
    {
        return strtolower((string) preg_replace('/(?<!^)[A-Z]/', '_$0', $key));
    }

    /** @return array<string,string> */
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
            0 => 'General', 1 => '0', 2 => '0.00', 3 => '#,##0', 4 => '#,##0.00',
            9 => '0%', 10 => '0.00%', 11 => '0.00E+00', 12 => '# ?/?', 13 => '# ??/??',
            14 => 'm/d/yy', 15 => 'd-mmm-yy', 16 => 'd-mmm', 17 => 'mmm-yy',
            18 => 'h:mm AM/PM', 19 => 'h:mm:ss AM/PM', 20 => 'h:mm', 21 => 'h:mm:ss',
            22 => 'm/d/yy h:mm', 37 => '#,##0 ;(#,##0)', 38 => '#,##0 ;[Red](#,##0)',
            39 => '#,##0.00;(#,##0.00)', 40 => '#,##0.00;[Red](#,##0.00)',
            45 => 'mm:ss', 46 => '[h]:mm:ss', 47 => 'mmss.0', 49 => '@',
            default => null,
        };
    }

    public static function looksLikeDateFormat(string $formatCode): bool
    {
        $format = strtolower($formatCode);
        $format = preg_replace('/"[^"]*"/', '', $format) ?? $format;
        $format = preg_replace('/\[[^\]]*\]/', '', $format) ?? $format;
        $format = str_replace(['\\', '_', '*'], '', $format);
        return preg_match('/(^|[^a-z])(y+|m+|d+|h+|s+|am\/pm)([^a-z]|$)/', $format) === 1
            && preg_match('/[ymdhs]/', $format) === 1;
    }
}
