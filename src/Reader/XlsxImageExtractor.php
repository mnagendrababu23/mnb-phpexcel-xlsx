<?php

declare(strict_types=1);

namespace Mnb\PHPExcel\Reader;

use Mnb\PHPExcel\Support\Coordinate;
use Mnb\PHPExcel\Support\MnbExcelException;
use Mnb\PHPExcel\Support\Zip\ZipArchive;

final class XlsxImageExtractor
{
    public function __construct(private readonly ?XlsxWorkbookResolver $resolver = null)
    {
    }

    /**
     * @return list<array<string,mixed>>
     */
    public function images(string $path, int|string $sheet = 1, bool $includeBytes = false): array
    {
        if (!class_exists(ZipArchive::class)) {
            throw new MnbExcelException('ext-zip is required to read XLSX images.');
        }
        $realPath = realpath($path);
        if ($realPath === false) {
            throw new MnbExcelException('Invalid XLSX path: ' . $path);
        }

        $resolver = $this->resolver ?? new XlsxWorkbookResolver();
        $sheetPath = $resolver->resolveSheetPath($realPath, $sheet);
        $zip = new ZipArchive();
        if ($zip->open($realPath) !== true) {
            throw new MnbExcelException('Unable to open XLSX zip package.');
        }

        try {
            $sheetXml = $zip->getFromName($sheetPath);
            if ($sheetXml === false) {
                return [];
            }
            $sheetRels = $this->relationships($zip, $this->relsPathForPart($sheetPath), dirname($sheetPath));
            $drawingPaths = [];
            preg_match_all('/<(?:[A-Za-z_][A-Za-z0-9_.-]*:)?(?:drawing|legacyDrawing|picture)\b[^>]*(?:[A-Za-z_][A-Za-z0-9_.-]*:)?id\s*=\s*("([^"]+)"|\'([^\']+)\')/i', $sheetXml, $drawingMatches, PREG_SET_ORDER);
            foreach ($drawingMatches as $match) {
                $rid = $match[2] !== '' ? $match[2] : $match[3];
                $relationship = $sheetRels[$rid] ?? null;
                if (is_array($relationship) && str_contains(strtolower((string) $relationship['type']), 'drawing') && $relationship['path'] !== null) {
                    $drawingPaths[] = (string) $relationship['path'];
                }
            }

            $images = [];
            foreach (array_values(array_unique($drawingPaths)) as $drawingPath) {
                $drawingXml = $zip->getFromName($drawingPath);
                if ($drawingXml === false) {
                    continue;
                }
                $drawingRels = $this->relationships($zip, $this->relsPathForPart($drawingPath), dirname($drawingPath));
                preg_match_all('/<(?:[A-Za-z_][A-Za-z0-9_.-]*:)?(oneCellAnchor|twoCellAnchor|absoluteAnchor)\b[^>]*>(.*?)<\/(?:[A-Za-z_][A-Za-z0-9_.-]*:)?\1>/is', $drawingXml, $anchors, PREG_SET_ORDER);
                foreach ($anchors as $anchorIndex => $anchorMatch) {
                    $anchorXml = $anchorMatch[2];
                    if (preg_match('/<(?:[A-Za-z_][A-Za-z0-9_.-]*:)?blip\b[^>]*(?:[A-Za-z_][A-Za-z0-9_.-]*:)?embed\s*=\s*("([^"]+)"|\'([^\']+)\')/i', $anchorXml, $blip) !== 1) {
                        continue;
                    }
                    $rid = $blip[2] !== '' ? $blip[2] : $blip[3];
                    $relationship = $drawingRels[$rid] ?? null;
                    if (!is_array($relationship) || $relationship['path'] === null) {
                        continue;
                    }
                    $mediaPath = (string) $relationship['path'];
                    $bytes = $zip->getFromName($mediaPath);
                    if ($bytes === false) {
                        continue;
                    }
                    $name = basename($mediaPath);
                    if (preg_match('/<(?:[A-Za-z_][A-Za-z0-9_.-]*:)?cNvPr\b([^>]*)\/?\s*>/i', $anchorXml, $nv) === 1) {
                        $nvAttrs = $this->parseAttributes('<xdr:cNvPr ' . $nv[1] . '>');
                        $name = (string) ($nvAttrs['name'] ?? $name);
                        $description = (string) ($nvAttrs['descr'] ?? '');
                    } else {
                        $description = '';
                    }
                    $from = $this->anchorMarker($anchorXml, 'from');
                    $to = $this->anchorMarker($anchorXml, 'to');
                    $extension = strtolower(pathinfo($mediaPath, PATHINFO_EXTENSION));
                    $image = [
                        'index' => count($images) + 1,
                        'name' => $name,
                        'description' => $description,
                        'media_path' => $mediaPath,
                        'drawing_path' => $drawingPath,
                        'relationship_id' => $rid,
                        'mime_type' => $this->mimeType($extension),
                        'extension' => $extension,
                        'size_bytes' => strlen($bytes),
                        'from' => $from,
                        'to' => $to,
                        'cell' => $from['cell'] ?? null,
                    ];
                    $size = @getimagesizefromstring($bytes);
                    if (is_array($size)) {
                        $image['width'] = (int) ($size[0] ?? 0);
                        $image['height'] = (int) ($size[1] ?? 0);
                    }
                    if ($includeBytes) {
                        $image['bytes'] = $bytes;
                    }
                    $images[] = $image;
                }
            }
            return $images;
        } finally {
            $zip->close();
        }
    }

    /** @return list<array<string,mixed>> */
    public function extract(string $path, string $directory, int|string $sheet = 1, bool $overwrite = false): array
    {
        if (!is_dir($directory) && !mkdir($directory, 0775, true) && !is_dir($directory)) {
            throw new MnbExcelException('Unable to create image extraction directory: ' . $directory);
        }

        $results = [];
        foreach ($this->images($path, $sheet, true) as $image) {
            $safeName = preg_replace('/[^A-Za-z0-9._-]+/', '_', (string) $image['name']) ?: ('image_' . $image['index'] . '.' . $image['extension']);
            if (pathinfo($safeName, PATHINFO_EXTENSION) === '') {
                $safeName .= '.' . $image['extension'];
            }
            $target = rtrim($directory, DIRECTORY_SEPARATOR) . DIRECTORY_SEPARATOR . $safeName;
            if (!$overwrite && is_file($target)) {
                $base = pathinfo($safeName, PATHINFO_FILENAME);
                $ext = pathinfo($safeName, PATHINFO_EXTENSION);
                $counter = 2;
                do {
                    $target = rtrim($directory, DIRECTORY_SEPARATOR) . DIRECTORY_SEPARATOR . $base . '_' . $counter . ($ext !== '' ? '.' . $ext : '');
                    $counter++;
                } while (is_file($target));
            }
            $bytes = (string) ($image['bytes'] ?? '');
            if (file_put_contents($target, $bytes) === false) {
                throw new MnbExcelException('Unable to extract XLSX image: ' . $target);
            }
            unset($image['bytes']);
            $image['path'] = $target;
            $results[] = $image;
        }
        return $results;
    }

    /** @return array<string,array<string,mixed>> */
    private function relationships(ZipArchive $zip, string $relsPath, string $baseDir): array
    {
        $xml = $zip->getFromName($relsPath);
        if ($xml === false) {
            return [];
        }
        $relationships = [];
        preg_match_all('/<(?:[A-Za-z_][A-Za-z0-9_.-]*:)?Relationship\b[^>]*\/?\s*>/i', $xml, $matches);
        foreach ($matches[0] ?? [] as $tag) {
            $attrs = $this->parseAttributes($tag);
            $id = (string) ($attrs['Id'] ?? '');
            if ($id === '') {
                continue;
            }
            $target = (string) ($attrs['Target'] ?? '');
            $external = strcasecmp((string) ($attrs['TargetMode'] ?? ''), 'External') === 0;
            $relationships[$id] = [
                'id' => $id,
                'type' => (string) ($attrs['Type'] ?? ''),
                'target' => $target,
                'path' => $external ? null : $this->resolveTargetPath($baseDir, $target),
            ];
        }
        return $relationships;
    }

    /** @return array<string,mixed> */
    private function anchorMarker(string $xml, string $tag): array
    {
        if (preg_match('/<(?:[A-Za-z_][A-Za-z0-9_.-]*:)?' . $tag . '\b[^>]*>(.*?)<\/(?:[A-Za-z_][A-Za-z0-9_.-]*:)?' . $tag . '>/is', $xml, $match) !== 1) {
            return [];
        }
        $values = [];
        foreach (['col', 'colOff', 'row', 'rowOff'] as $name) {
            if (preg_match('/<(?:[A-Za-z_][A-Za-z0-9_.-]*:)?' . $name . '\b[^>]*>(-?\d+)<\/(?:[A-Za-z_][A-Za-z0-9_.-]*:)?' . $name . '>/i', $match[1], $valueMatch) === 1) {
                $values[$name === 'colOff' ? 'column_offset_emu' : ($name === 'rowOff' ? 'row_offset_emu' : $name)] = (int) $valueMatch[1];
            }
        }
        if (isset($values['col'], $values['row'])) {
            $values['column'] = $values['col'] + 1;
            $values['row'] = $values['row'] + 1;
            $values['cell'] = Coordinate::columnIndexToName($values['column']) . $values['row'];
            unset($values['col']);
        }
        return $values;
    }

    /** @return array<string,string> */
    private function parseAttributes(string $tag): array
    {
        $attrs = [];
        preg_match_all('/([A-Za-z0-9_:\-]+)\s*=\s*("([^"]*)"|\'([^\']*)\')/u', $tag, $matches, PREG_SET_ORDER);
        foreach ($matches as $match) {
            $attrs[$match[1]] = html_entity_decode($match[3] !== '' ? $match[3] : $match[4], ENT_QUOTES | ENT_XML1, 'UTF-8');
        }
        return $attrs;
    }

    private function relsPathForPart(string $partPath): string
    {
        $dir = dirname($partPath);
        return ($dir === '.' ? '_rels/' : $dir . '/_rels/') . basename($partPath) . '.rels';
    }

    private function resolveTargetPath(string $baseDir, string $target): string
    {
        $candidate = str_starts_with($target, '/') ? ltrim($target, '/') : rtrim($baseDir, '/') . '/' . $target;
        $parts = [];
        foreach (explode('/', str_replace('\\', '/', $candidate)) as $part) {
            if ($part === '' || $part === '.') {
                continue;
            }
            if ($part === '..') {
                array_pop($parts);
            } else {
                $parts[] = $part;
            }
        }
        return implode('/', $parts);
    }

    private function mimeType(string $extension): string
    {
        return match ($extension) {
            'png' => 'image/png',
            'jpg', 'jpeg' => 'image/jpeg',
            'gif' => 'image/gif',
            'bmp' => 'image/bmp',
            'tif', 'tiff' => 'image/tiff',
            'svg' => 'image/svg+xml',
            'emf' => 'image/emf',
            'wmf' => 'image/wmf',
            default => 'application/octet-stream',
        };
    }
}
