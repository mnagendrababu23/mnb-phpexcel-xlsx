<?php

declare(strict_types=1);

namespace Mnb\PHPExcel\Reader;

use Mnb\PHPExcel\Core\RichText;
use Mnb\PHPExcel\Support\MnbExcelException;
use ZipArchive;

final class XlsxMetadataExtractor
{
    private XlsxWorkbookResolver $resolver;

    public function __construct(?XlsxWorkbookResolver $resolver = null)
    {
        $this->resolver = $resolver ?? new XlsxWorkbookResolver();
    }

    /**
     * Extract rich XLSX metadata that is intentionally separate from plain cell values.
     *
     * @return array{
     *   rich_text:list<array<string,mixed>>,
     *   comments:list<array<string,mixed>>,
     *   hyperlinks:list<array<string,mixed>>,
     *   advanced_objects:array<string,mixed>,
     *   summary:array<string,int>
     * }
     */
    public function readSheetMetadata(string $path, int|string $sheet = 1): array
    {
        if (!class_exists(ZipArchive::class)) {
            throw new MnbExcelException('ext-zip is required to read XLSX metadata.');
        }

        $realPath = realpath($path);
        if ($realPath === false) {
            throw new MnbExcelException('Invalid XLSX path: ' . $path);
        }

        $sheetPath = $this->resolver->resolveSheetPath($realPath, $sheet);
        $zip = new ZipArchive();
        if ($zip->open($realPath) !== true) {
            throw new MnbExcelException('Unable to open XLSX zip package.');
        }

        $sheetXml = $zip->getFromName($sheetPath);
        if ($sheetXml === false) {
            $zip->close();
            throw new MnbExcelException('Unable to open worksheet XML: ' . $sheetPath);
        }

        $relationships = $this->readRelationships($zip, $this->relsPathForPart($sheetPath), dirname($sheetPath));
        $sharedStrings = $this->readSharedStringRichText($zip);
        $richText = $this->extractRichTextCells($sheetXml, $sharedStrings);
        $comments = $this->extractComments($zip, $relationships);
        $hyperlinks = $this->extractHyperlinks($sheetXml, $relationships);
        $advancedObjects = $this->extractAdvancedObjects($zip, $sheetXml, $relationships);

        $zip->close();
        $richTextObjects = [];
        foreach ($richText as $item) {
            $richTextObjects[(string) $item['cell']] = RichText::fromArray((array) ($item['runs'] ?? []));
        }
        $images = (new XlsxImageExtractor($this->resolver))->images($realPath, $sheet, false);

        return [
            'rich_text' => $richText,
            'rich_text_objects' => $richTextObjects,
            'comments' => $comments,
            'hyperlinks' => $hyperlinks,
            'images' => $images,
            'advanced_objects' => $advancedObjects,
            'summary' => [
                'rich_text_cells' => count($richText),
                'comments' => count($comments),
                'hyperlinks' => count($hyperlinks),
                'images' => count($images),
                'advanced_object_parts' => count((array) ($advancedObjects['package_parts'] ?? [])),
            ],
        ];
    }

    /** @return array<string,array{id:string,type:string,target:string,target_mode:string,path:?string}> */
    private function readRelationships(ZipArchive $zip, string $relsPath, string $baseDir): array
    {
        $xml = $zip->getFromName($relsPath);
        if ($xml === false || trim($xml) === '') {
            return [];
        }

        $relationships = [];
        foreach ($this->matchTags($xml, 'Relationship') as $tag) {
            $attrs = $this->parseAttributes($tag);
            $id = (string) ($attrs['Id'] ?? '');
            if ($id === '') {
                continue;
            }

            $target = (string) ($attrs['Target'] ?? '');
            $targetMode = (string) ($attrs['TargetMode'] ?? 'Internal');
            $relationships[$id] = [
                'id' => $id,
                'type' => (string) ($attrs['Type'] ?? ''),
                'target' => $target,
                'target_mode' => $targetMode,
                'path' => strcasecmp($targetMode, 'External') === 0 ? null : $this->resolveTargetPath($baseDir, $target),
            ];
        }

        return $relationships;
    }

    /** @return array<int,array{text:string,runs:list<array<string,mixed>>,rich:bool}> */
    private function readSharedStringRichText(ZipArchive $zip): array
    {
        $xml = $zip->getFromName('xl/sharedStrings.xml');
        if ($xml === false || trim($xml) === '') {
            return [];
        }

        $strings = [];
        preg_match_all('/<(?:[A-Za-z_][A-Za-z0-9_.-]*:)?si\b[^>]*>.*?<\/(?:[A-Za-z_][A-Za-z0-9_.-]*:)?si>/isu', $xml, $matches);
        foreach ($matches[0] ?? [] as $siXml) {
            $rich = $this->richTextFromXml($siXml);
            $strings[] = [
                'text' => $rich['text'],
                'runs' => $rich['runs'],
                'rich' => $this->hasRichTextFormatting($rich['runs']),
            ];
        }

        return $strings;
    }

    /**
     * @param array<int,array{text:string,runs:list<array<string,mixed>>,rich:bool}> $sharedStrings
     * @return list<array<string,mixed>>
     */
    private function extractRichTextCells(string $sheetXml, array $sharedStrings): array
    {
        $cells = [];
        preg_match_all('/<(?:[A-Za-z_][A-Za-z0-9_.-]*:)?c\b[^>]*(?:\/>|>.*?<\/(?:[A-Za-z_][A-Za-z0-9_.-]*:)?c>)/isu', $sheetXml, $matches);
        foreach ($matches[0] ?? [] as $cellXml) {
            $attrs = $this->parseAttributes($this->openingTag($cellXml));
            $cell = strtoupper((string) ($attrs['r'] ?? ''));
            if ($cell === '') {
                continue;
            }

            $type = (string) ($attrs['t'] ?? '');
            $rich = null;
            if ($type === 's') {
                $index = $this->readValueIndex($cellXml);
                if ($index !== null && isset($sharedStrings[$index]) && $sharedStrings[$index]['rich']) {
                    $rich = $sharedStrings[$index];
                    $rich['source'] = 'shared_string';
                    $rich['shared_string_index'] = $index;
                }
            } elseif ($type === 'inlineStr' && preg_match('/<(?:[A-Za-z_][A-Za-z0-9_.-]*:)?r\b/i', $cellXml) === 1) {
                $parsed = $this->richTextFromXml($cellXml);
                if ($this->hasRichTextFormatting($parsed['runs'])) {
                    $rich = $parsed + ['source' => 'inline_string'];
                }
            }

            if ($rich !== null) {
                $cells[] = [
                    'cell' => $cell,
                    'text' => (string) $rich['text'],
                    'runs' => array_values((array) $rich['runs']),
                    'source' => $rich['source'] ?? 'unknown',
                    'shared_string_index' => $rich['shared_string_index'] ?? null,
                ];
            }
        }

        return $cells;
    }

    /** @param array<string,array{id:string,type:string,target:string,target_mode:string,path:?string}> $relationships @return list<array<string,mixed>> */
    private function extractHyperlinks(string $sheetXml, array $relationships): array
    {
        $hyperlinks = [];
        if (!preg_match('/<(?:[A-Za-z_][A-Za-z0-9_.-]*:)?hyperlinks\b[^>]*>(.*?)<\/(?:[A-Za-z_][A-Za-z0-9_.-]*:)?hyperlinks>/isu', $sheetXml, $container)) {
            return [];
        }

        foreach ($this->matchTags($container[1], 'hyperlink') as $tag) {
            $attrs = $this->parseAttributes($tag);
            $ref = strtoupper((string) ($attrs['ref'] ?? ''));
            if ($ref === '') {
                continue;
            }

            $relationshipId = (string) ($attrs['r:id'] ?? $attrs['id'] ?? '');
            $relationship = $relationshipId !== '' ? ($relationships[$relationshipId] ?? null) : null;
            $target = $relationship['target'] ?? null;
            $targetMode = $relationship['target_mode'] ?? null;

            $hyperlinks[] = [
                'cell' => $ref,
                'ref' => $ref,
                'target' => $target,
                'location' => $attrs['location'] ?? null,
                'display' => $attrs['display'] ?? null,
                'tooltip' => $attrs['tooltip'] ?? null,
                'relationship_id' => $relationshipId !== '' ? $relationshipId : null,
                'target_mode' => $targetMode,
                'type' => $targetMode !== null && strcasecmp((string) $targetMode, 'External') === 0 ? 'external' : 'internal',
            ];
        }

        return $hyperlinks;
    }

    /** @param array<string,array{id:string,type:string,target:string,target_mode:string,path:?string}> $relationships @return list<array<string,mixed>> */
    private function extractComments(ZipArchive $zip, array $relationships): array
    {
        $comments = [];
        foreach ($relationships as $relationship) {
            $type = strtolower($relationship['type']);
            $path = $relationship['path'];
            if ($path === null || (!str_contains($type, '/comments') && !str_contains($type, '/threadedcomments'))) {
                continue;
            }

            $xml = $zip->getFromName($path);
            if ($xml === false || trim($xml) === '') {
                continue;
            }

            if (str_contains($type, '/threadedcomments')) {
                $comments = array_merge($comments, $this->extractThreadedComments($xml, $path));
                continue;
            }

            $authors = [];
            if (preg_match('/<(?:[A-Za-z_][A-Za-z0-9_.-]*:)?authors\b[^>]*>(.*?)<\/(?:[A-Za-z_][A-Za-z0-9_.-]*:)?authors>/isu', $xml, $authorContainer)) {
                preg_match_all('/<(?:[A-Za-z_][A-Za-z0-9_.-]*:)?author\b[^>]*>(.*?)<\/(?:[A-Za-z_][A-Za-z0-9_.-]*:)?author>/isu', $authorContainer[1], $authorMatches);
                foreach ($authorMatches[1] ?? [] as $author) {
                    $authors[] = $this->decode($this->stripTags($author));
                }
            }

            preg_match_all('/<(?:[A-Za-z_][A-Za-z0-9_.-]*:)?comment\b[^>]*(?:\/>|>.*?<\/(?:[A-Za-z_][A-Za-z0-9_.-]*:)?comment>)/isu', $xml, $commentMatches);
            foreach ($commentMatches[0] ?? [] as $commentXml) {
                $attrs = $this->parseAttributes($this->openingTag($commentXml));
                $cell = strtoupper((string) ($attrs['ref'] ?? ''));
                if ($cell === '') {
                    continue;
                }
                $authorId = isset($attrs['authorId']) && ctype_digit((string) $attrs['authorId']) ? (int) $attrs['authorId'] : null;
                $rich = $this->richTextFromXml($commentXml);
                $comments[] = [
                    'cell' => $cell,
                    'author' => $authorId !== null ? ($authors[$authorId] ?? null) : null,
                    'author_id' => $authorId,
                    'text' => $rich['text'],
                    'runs' => $rich['runs'],
                    'part' => $path,
                    'type' => 'comment',
                ];
            }
        }

        return $comments;
    }

    /** @return list<array<string,mixed>> */
    private function extractThreadedComments(string $xml, string $path): array
    {
        $comments = [];
        preg_match_all('/<(?:[A-Za-z_][A-Za-z0-9_.-]*:)?threadedComment\b[^>]*(?:\/>|>.*?<\/(?:[A-Za-z_][A-Za-z0-9_.-]*:)?threadedComment>)/isu', $xml, $matches);
        foreach ($matches[0] ?? [] as $commentXml) {
            $attrs = $this->parseAttributes($this->openingTag($commentXml));
            $cell = strtoupper((string) ($attrs['ref'] ?? ''));
            if ($cell === '') {
                continue;
            }
            $rich = $this->richTextFromXml($commentXml);
            $comments[] = [
                'cell' => $cell,
                'author' => null,
                'author_id' => $attrs['personId'] ?? null,
                'text' => $rich['text'],
                'runs' => $rich['runs'],
                'part' => $path,
                'type' => 'threaded_comment',
            ];
        }

        return $comments;
    }

    /** @param array<string,array{id:string,type:string,target:string,target_mode:string,path:?string}> $relationships @return array<string,mixed> */
    private function extractAdvancedObjects(ZipArchive $zip, string $sheetXml, array $relationships): array
    {
        $relationshipParts = [];
        foreach ($relationships as $relationship) {
            $type = strtolower($relationship['type']);
            if (!$this->isAdvancedRelationshipType($type)) {
                continue;
            }

            $relationshipParts[] = [
                'relationship_id' => $relationship['id'],
                'type' => $relationship['type'],
                'target' => $relationship['target'],
                'target_mode' => $relationship['target_mode'],
                'path' => $relationship['path'],
            ];
        }

        $sheetElements = [];
        foreach (['drawing', 'legacyDrawing', 'legacyDrawingHF', 'picture', 'tableParts', 'oleObjects', 'controls', 'webPublishItems', 'extLst'] as $tag) {
            if (preg_match('/<(?:[A-Za-z_][A-Za-z0-9_.-]*:)?' . preg_quote($tag, '/') . '\b/i', $sheetXml) === 1) {
                $sheetElements[] = $tag;
            }
        }

        $packageParts = [];
        for ($i = 0; $i < $zip->numFiles; $i++) {
            $name = (string) $zip->getNameIndex($i);
            if ($this->isAdvancedPackagePart($name)) {
                $packageParts[] = $name;
            }
        }

        sort($packageParts);

        return [
            'sheet_elements' => $sheetElements,
            'relationships' => $relationshipParts,
            'package_parts' => $packageParts,
            'preservation_supported' => true,
            'preservation_note' => 'Use WorkbookBuilder::preserveAdvancedObjectsFrom() when rewriting from a source XLSX template.',
        ];
    }

    /** @return array{text:string,runs:list<array<string,mixed>>} */
    private function richTextFromXml(string $xml): array
    {
        $runs = [];
        preg_match_all('/<(?:[A-Za-z_][A-Za-z0-9_.-]*:)?r\b[^>]*>.*?<\/(?:[A-Za-z_][A-Za-z0-9_.-]*:)?r>/isu', $xml, $runMatches);
        foreach ($runMatches[0] ?? [] as $runXml) {
            $text = $this->textFromTElements($runXml);
            if ($text === '') {
                continue;
            }
            $runs[] = ['text' => $text] + $this->runProperties($runXml);
        }

        if ($runs === []) {
            $text = $this->textFromTElements($xml);
            if ($text !== '') {
                $runs[] = ['text' => $text];
            }
        }

        $plain = '';
        foreach ($runs as $run) {
            $plain .= (string) $run['text'];
        }

        return ['text' => $plain, 'runs' => $runs];
    }

    /** @return array<string,mixed> */
    private function runProperties(string $runXml): array
    {
        if (!preg_match('/<(?:[A-Za-z_][A-Za-z0-9_.-]*:)?rPr\b[^>]*>(.*?)<\/(?:[A-Za-z_][A-Za-z0-9_.-]*:)?rPr>/isu', $runXml, $match)) {
            return [];
        }

        $xml = $match[1];
        $props = [];
        foreach (['b' => 'bold', 'i' => 'italic', 'strike' => 'strike', 'outline' => 'outline', 'shadow' => 'shadow', 'condense' => 'condense', 'extend' => 'extend'] as $tag => $key) {
            if (preg_match('/<(?:[A-Za-z_][A-Za-z0-9_.-]*:)?' . $tag . '\b([^>]*)\/?\s*>/iu', $xml, $m) === 1) {
                $attrs = $this->parseAttributes('<' . $tag . ' ' . ($m[1] ?? '') . '>');
                $value = strtolower((string) ($attrs['val'] ?? '1'));
                $props[$key] = !in_array($value, ['0', 'false', 'off', 'no'], true);
            }
        }
        if (preg_match('/<(?:[A-Za-z_][A-Za-z0-9_.-]*:)?u\b([^>]*)\/>/iu', $xml, $m) === 1) {
            $attrs = $this->parseAttributes('<u' . $m[1] . '/>');
            $props['underline'] = $attrs['val'] ?? true;
        }
        if (preg_match('/<(?:[A-Za-z_][A-Za-z0-9_.-]*:)?rFont\b[^>]*val\s*=\s*("([^"]*)"|\'([^\']*)\')/iu', $xml, $m) === 1) {
            $props['font'] = $this->decode($m[2] !== '' ? $m[2] : $m[3]);
        }
        if (preg_match('/<(?:[A-Za-z_][A-Za-z0-9_.-]*:)?sz\b[^>]*val\s*=\s*("([^"]*)"|\'([^\']*)\')/iu', $xml, $m) === 1) {
            $props['size'] = (float) ($m[2] !== '' ? $m[2] : $m[3]);
        }
        if (preg_match('/<(?:[A-Za-z_][A-Za-z0-9_.-]*:)?color\b([^>]*)\/>/iu', $xml, $m) === 1) {
            $attrs = $this->parseAttributes('<color' . $m[1] . '/>');
            foreach (['rgb', 'indexed', 'theme', 'tint'] as $key) {
                if (isset($attrs[$key])) {
                    $props['color'][$key] = $attrs[$key];
                }
            }
        }
        if (preg_match('/<(?:[A-Za-z_][A-Za-z0-9_.-]*:)?vertAlign\b[^>]*val\s*=\s*("([^"]*)"|\'([^\']*)\')/iu', $xml, $m) === 1) {
            $props['vertical_align'] = $this->decode($m[2] !== '' ? $m[2] : $m[3]);
        }
        foreach (['family' => 'family', 'charset' => 'charset', 'scheme' => 'scheme'] as $tag => $key) {
            if (preg_match('/<(?:[A-Za-z_][A-Za-z0-9_.-]*:)?' . $tag . '\b[^>]*val\s*=\s*("([^"]*)"|\'([^\']*)\')/iu', $xml, $m) === 1) {
                $value = $this->decode($m[2] !== '' ? $m[2] : $m[3]);
                $props[$key] = is_numeric($value) ? (int) $value : $value;
            }
        }

        return $props;
    }

    /** @param list<array<string,mixed>> $runs */
    private function hasRichTextFormatting(array $runs): bool
    {
        if (count($runs) > 1) {
            return true;
        }
        foreach ($runs as $run) {
            if (count($run) > 1) {
                return true;
            }
        }
        return false;
    }

    private function textFromTElements(string $xml): string
    {
        preg_match_all('/<(?:[A-Za-z_][A-Za-z0-9_.-]*:)?t\b[^>]*>(.*?)<\/(?:[A-Za-z_][A-Za-z0-9_.-]*:)?t>/isu', $xml, $matches);
        $text = '';
        foreach ($matches[1] ?? [] as $part) {
            $text .= $this->decode($this->stripTags($part));
        }
        return $text;
    }

    /** @return list<string> */
    private function matchTags(string $xml, string $tag): array
    {
        preg_match_all('/<(?:[A-Za-z_][A-Za-z0-9_.-]*:)?' . preg_quote($tag, '/') . '\b[^>]*(?:\/>|>.*?<\/(?:[A-Za-z_][A-Za-z0-9_.-]*:)?' . preg_quote($tag, '/') . '>)/isu', $xml, $matches);
        return $matches[0] ?? [];
    }

    /** @return array<string,string> */
    private function parseAttributes(string $tag): array
    {
        $attrs = [];
        preg_match_all('/([A-Za-z0-9_:\-]+)\s*=\s*("([^"]*)"|\'([^\']*)\')/u', $tag, $matches, PREG_SET_ORDER);
        foreach ($matches as $match) {
            $attrs[$match[1]] = $this->decode($match[3] !== '' ? $match[3] : $match[4]);
        }
        return $attrs;
    }

    private function openingTag(string $xml): string
    {
        if (preg_match('/^<[^>]+>/s', trim($xml), $match) === 1) {
            return $match[0];
        }
        return '';
    }

    private function readValueIndex(string $cellXml): ?int
    {
        if (preg_match('/<(?:[A-Za-z_][A-Za-z0-9_.-]*:)?v\b[^>]*>(.*?)<\/(?:[A-Za-z_][A-Za-z0-9_.-]*:)?v>/isu', $cellXml, $match) !== 1) {
            return null;
        }
        $value = trim($this->decode($this->stripTags($match[1])));
        return ctype_digit($value) ? (int) $value : null;
    }

    private function relsPathForPart(string $partPath): string
    {
        $dir = dirname($partPath);
        $base = basename($partPath);
        return ($dir === '.' ? '_rels/' : $dir . '/_rels/') . $base . '.rels';
    }

    private function resolveTargetPath(string $baseDir, string $target): string
    {
        $target = str_replace('\\', '/', trim($target));
        if ($target === '') {
            return $this->normalizePath($baseDir);
        }
        if (str_starts_with($target, '/')) {
            return $this->normalizePath(ltrim($target, '/'));
        }
        return $this->normalizePath(rtrim($baseDir, '/') . '/' . $target);
    }

    private function normalizePath(string $path): string
    {
        $parts = [];
        foreach (explode('/', str_replace('\\', '/', $path)) as $part) {
            if ($part === '' || $part === '.') {
                continue;
            }
            if ($part === '..') {
                array_pop($parts);
                continue;
            }
            $parts[] = $part;
        }
        return implode('/', $parts);
    }

    private function stripTags(string $value): string
    {
        return preg_replace('/<[^>]+>/', '', $value) ?? '';
    }

    private function decode(string $value): string
    {
        return html_entity_decode($value, ENT_QUOTES | ENT_XML1, 'UTF-8');
    }

    private function isAdvancedRelationshipType(string $type): bool
    {
        foreach (['drawing', 'vmldrawing', 'comments', 'threadedcomments', 'table', 'pivot', 'chart', 'oleobject', 'ctrlprop', 'externalLink', 'printersettings', 'querytable', 'slicer'] as $needle) {
            if (str_contains($type, strtolower($needle))) {
                return true;
            }
        }
        return false;
    }

    private function isAdvancedPackagePart(string $name): bool
    {
        $lower = strtolower($name);
        foreach ([
            'xl/drawings/', 'xl/media/', 'xl/comments', 'xl/threadedcomments/', 'xl/persons/', 'xl/vmldrawings/',
            'xl/tables/', 'xl/pivottables/', 'xl/pivotcache/', 'xl/charts/', 'xl/embeddings/', 'xl/ctrlprops/',
            'xl/printersettings/', 'xl/externallinks/', 'xl/querytables/', 'xl/slicers/', 'xl/timelines/', 'xl/model/',
        ] as $prefix) {
            if (str_starts_with($lower, $prefix)) {
                return true;
            }
        }
        return false;
    }
}
