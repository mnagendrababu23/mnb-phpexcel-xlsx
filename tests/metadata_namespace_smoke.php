<?php

declare(strict_types=1);

require __DIR__ . '/../vendor/autoload.php';

use Mnb\PHPExcel\Core\WorkbookData;
use Mnb\PHPExcel\Core\WorksheetData;
use Mnb\PHPExcel\Format\Xlsx;
use Mnb\PHPExcel\Support\XlsxIntegrityValidator;
use Mnb\PHPExcel\Support\Zip\ZipArchive;
use Mnb\PHPExcel\Writer\XlsxWriter;

function check(bool $ok, string $message): void { if (!$ok) throw new RuntimeException($message); }
function prefix_default_document(string $xml, string $prefix, string $uri): string {
    $xml = preg_replace('/\\sxmlns="' . preg_quote($uri, '/') . '"/', ' xmlns:' . $prefix . '="' . $uri . '"', $xml, 1) ?? $xml;
    return preg_replace_callback('/<(\\/?)([A-Za-z_][A-Za-z0-9_.-]*)(?=[\\s>\\/])/', static function (array $m) use ($prefix): string {
        if (str_contains($m[2], ':')) return $m[0];
        return '<' . $m[1] . $prefix . ':' . $m[2];
    }, $xml) ?? $xml;
}

$dir = '/tmp/mnb-metadata-ns-' . bin2hex(random_bytes(4));
mkdir($dir, 0775, true);
$source = $dir . '/prefixed.xlsx';
$dest = $dir . '/updated.xlsx';
try {
    (new XlsxWriter())->write(new WorkbookData([
        new WorksheetData('Data', [['Name'], ['Alice']]),
        new WorksheetData('Archive', [['Value'], [1]]),
    ], ['title' => 'Original', 'creator' => 'MNB']), $source);

    $zip = new ZipArchive();
    check($zip->open($source) === true, 'open source');

    $workbook = (string) $zip->getFromName('xl/workbook.xml');
    $workbook = prefix_default_document($workbook, 'x', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main');
    $zip->addFromString('xl/workbook.xml', $workbook);

    $core = (string) $zip->getFromName('docProps/core.xml');
    $core = str_replace(
        ['xmlns:cp=', 'xmlns:dc=', 'xmlns:dcterms=', 'xmlns:dcmitype=', 'xmlns:xsi=', 'cp:', 'dc:', 'dcterms:', 'dcmitype:', 'xsi:'],
        ['xmlns:c=', 'xmlns:d=', 'xmlns:t=', 'xmlns:m=', 'xmlns:i=', 'c:', 'd:', 't:', 'm:', 'i:'],
        $core
    );
    $zip->addFromString('docProps/core.xml', $core);

    $app = (string) $zip->getFromName('docProps/app.xml');
    $app = prefix_default_document($app, 'e', 'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties');
    $zip->addFromString('docProps/app.xml', $app);

    $types = (string) $zip->getFromName('[Content_Types].xml');
    $types = prefix_default_document($types, 'ct', 'http://schemas.openxmlformats.org/package/2006/content-types');
    $zip->addFromString('[Content_Types].xml', $types);

    $rels = (string) $zip->getFromName('_rels/.rels');
    $rels = prefix_default_document($rels, 'pr', 'http://schemas.openxmlformats.org/package/2006/relationships');
    $zip->addFromString('_rels/.rels', $rels);
    $zip->close();

    (new XlsxIntegrityValidator())->assertValid($source);

    Xlsx::updateMetaInfo($source, $dest, [
        'document' => ['title' => 'Prefixed Updated', 'identifier' => 'DOC-42'],
        'revision' => ['document_created_at' => '2026-08-02T10:00:00Z'],
        'application' => ['company' => 'MNB Labs', 'manager' => 'Engineering'],
        'custom_properties' => ['Project ID' => 'NS-1'],
        'workbook' => [
            'sheet_visibility' => ['Data' => 'hidden', 'Archive' => 'visible'],
            'active_sheet' => 'Archive',
            'date1904' => true,
        ],
        'calculation' => ['mode' => 'manual'],
    ]);

    (new XlsxIntegrityValidator())->assertValid($dest);
    $meta = Xlsx::metaInfo($dest, ['profile' => 'full']);
    check($meta['document']['title'] === 'Prefixed Updated', 'title');
    check($meta['document']['identifier'] === 'DOC-42', 'identifier');
    check($meta['application']['company'] === 'MNB Labs', 'company');
    check($meta['application']['manager'] === 'Engineering', 'manager');
    check($meta['workbook']['active_sheet']['name'] === 'Archive', 'active');
    check($meta['workbook']['date1904'] === true, 'date1904');
    check($meta['calculation']['settings']['calcMode'] === 'manual', 'calc');
    check($meta['custom_properties']['count'] === 1, 'custom');

    $zip = new ZipArchive();
    check($zip->open($dest) === true, 'open dest');
    $workbook = (string) $zip->getFromName('xl/workbook.xml');
    $core = (string) $zip->getFromName('docProps/core.xml');
    $app = (string) $zip->getFromName('docProps/app.xml');
    $types = (string) $zip->getFromName('[Content_Types].xml');
    $rels = (string) $zip->getFromName('_rels/.rels');
    $zip->close();
    check(str_contains($workbook, '<x:calcPr'), 'workbook prefix preserved for inserted calcPr');
    check(str_contains($core, '<d:identifier>DOC-42</d:identifier>'), 'core namespace alias used');
    check(str_contains($app, '<e:Manager>Engineering</e:Manager>'), 'app namespace alias used');
    check(str_contains($types, '<ct:Override'), 'content type prefix used');
    check(str_contains($rels, '<pr:Relationship'), 'relationship prefix used');

    echo "metadata_namespace_smoke passed\n";
} finally {
    @unlink($source); @unlink($dest); @rmdir($dir);
}
