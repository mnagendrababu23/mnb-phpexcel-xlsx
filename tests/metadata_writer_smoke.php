<?php

declare(strict_types=1);

require __DIR__ . '/../vendor/autoload.php';

use Mnb\PHPExcel\Core\WorkbookData;
use Mnb\PHPExcel\Core\WorksheetData;
use Mnb\PHPExcel\Format\Xlsx;
use Mnb\PHPExcel\Support\Zip\ZipArchive;
use Mnb\PHPExcel\Writer\XlsxWriter;

function metadata_writer_check(bool $condition, string $message): void
{
    if (!$condition) {
        throw new RuntimeException($message);
    }
}

$directory = sys_get_temp_dir() . '/mnb-xlsx-metadata-writer-' . bin2hex(random_bytes(5));
if (!mkdir($directory, 0775, true) && !is_dir($directory)) {
    throw new RuntimeException('Unable to create metadata writer smoke-test directory.');
}
$source = $directory . '/source.xlsx';
$updated = $directory . '/updated.xlsx';
$clean = $directory . '/clean.xlsx';
$encrypted = $directory . '/encrypted.xlsx';
$encryptedUpdated = $directory . '/encrypted-updated.xlsx';
$invalid = $directory . '/invalid.xlsx';
$wrongPassword = $directory . '/wrong-password.xlsx';
$standardEncrypted = $directory . '/standard-encrypted.xlsx';
$standardUpdated = $directory . '/standard-updated.xlsx';
$password = 'Mnb-Metadata-Test-Password!';

try {
    (new XlsxWriter())->write(new WorkbookData([
        new WorksheetData('Data', [['Name'], ['Alice']], comments: [
            ['cell' => 'A2', 'author' => 'Reviewer', 'text' => 'Review this value.'],
        ]),
        new WorksheetData('Archive', [['Value'], [1]]),
    ], [
        'title' => 'Original Title',
        'creator' => 'Original Author',
        'last_modified_by' => 'Original Editor',
        'company' => 'Original Company',
    ]), $source);

    // Add a valid unknown OOXML part. Metadata updates must preserve it byte-for-byte.
    $zip = new ZipArchive();
    metadata_writer_check($zip->open($source) === true, 'Unable to open source package.');
    $opaque = 'opaque-data-' . bin2hex(random_bytes(8));
    $zip->addFromString('customXml/opaque.bin', $opaque);
    $contentTypes = (string) $zip->getFromName('[Content_Types].xml');
    if (!preg_match('/<Default\s+Extension=["\']bin["\']/i', $contentTypes)) {
        $contentTypes = preg_replace(
            '/<\/Types>\s*$/',
            '<Default Extension="bin" ContentType="application/octet-stream"/></Types>',
            $contentTypes,
            1,
        ) ?? $contentTypes;
        $zip->addFromString('[Content_Types].xml', $contentTypes);
    }
    $zip->close();
    $opaqueHash = hash('sha256', $opaque);

    Xlsx::updateMetaInfo($source, $updated, [
        'document' => [
            'title' => 'Updated Title',
            'creator' => 'New Author',
            'keywords' => ['metadata', 'xlsx'],
        ],
        'revision' => [
            'last_saved_by' => 'Release Bot',
            'revision_number' => '7',
            'total_editing_time_seconds' => 125,
            'document_created_at' => '2026-01-01T10:00:00+05:30',
        ],
        'application' => [
            'application_version' => '2.1.0',
            'manager' => 'Engineering Manager',
            'company' => 'MNB Labs',
        ],
        'custom_properties' => [
            'Project ID' => 'PRJ-2002',
            'Approved' => true,
        ],
        'workbook' => [
            'sheet_visibility' => ['Data' => 'hidden', 'Archive' => 'visible'],
            'active_sheet' => 'Archive',
            'date1904' => true,
            'code_name' => 'MetadataBook',
            'read_only_recommended' => true,
        ],
        'calculation' => [
            'mode' => 'manual',
            'calc_id' => 20260802,
            'calc_on_save' => false,
        ],
    ]);

    $sourceInfo = Xlsx::metaInfo($source, ['profile' => 'quick']);
    $updatedInfo = Xlsx::metaInfo($updated, ['profile' => 'full']);
    metadata_writer_check($sourceInfo['document']['title'] === 'Original Title', 'Source workbook was changed.');
    metadata_writer_check($updatedInfo['document']['title'] === 'Updated Title', 'Document title was not updated.');
    metadata_writer_check($updatedInfo['revision']['last_saved_by'] === 'Release Bot', 'Revision author was not updated.');
    metadata_writer_check($updatedInfo['revision']['total_editing_time_seconds'] === 180, 'Editing time was not normalized to Office minutes.');
    metadata_writer_check($updatedInfo['application']['company'] === 'MNB Labs', 'Company was not updated.');
    metadata_writer_check($updatedInfo['workbook']['active_sheet']['name'] === 'Archive', 'Active sheet was not updated.');
    metadata_writer_check($updatedInfo['workbook']['date1904'] === true, 'Date system was not updated.');
    metadata_writer_check($updatedInfo['calculation']['settings']['calcMode'] === 'manual', 'Calculation mode was not updated.');

    $zip = new ZipArchive();
    metadata_writer_check($zip->open($updated) === true, 'Unable to open updated package.');
    $preserved = $zip->getFromName('customXml/opaque.bin');
    $zip->close();
    metadata_writer_check(is_string($preserved) && hash('sha256', $preserved) === $opaqueHash, 'Unknown package part was not preserved.');

    Xlsx::updateMetaInfo($updated, $updated, ['document' => ['subject' => 'Atomic in-place update']]);
    metadata_writer_check(
        Xlsx::metaInfo($updated, ['profile' => 'quick'])['document']['subject'] === 'Atomic in-place update',
        'Atomic same-path metadata update failed.'
    );

    Xlsx::removePersonalInfo($updated, $clean);
    $cleanInfo = Xlsx::metaInfo($clean, ['profile' => 'full']);
    metadata_writer_check($cleanInfo['document']['creator'] === null, 'Creator was not removed.');
    metadata_writer_check($cleanInfo['revision']['last_saved_by'] === null, 'Last-saved-by was not removed.');
    metadata_writer_check($cleanInfo['application']['manager'] === null, 'Manager was not removed.');
    metadata_writer_check($cleanInfo['application']['company'] === null, 'Company was not removed.');
    metadata_writer_check($cleanInfo['custom_properties']['count'] === 0, 'Custom properties were not removed.');
    metadata_writer_check($cleanInfo['comments_notes']['items'][0]['author'] === 'Author', 'Comment author was not anonymized.');

    $thrown = false;
    try {
        Xlsx::updateMetaInfo($source, $invalid, [
            'workbook' => ['sheet_visibility' => ['Data' => 'hidden', 'Archive' => 'hidden']],
        ]);
    } catch (Throwable) {
        $thrown = true;
    }
    metadata_writer_check($thrown, 'All-hidden worksheet update was not rejected.');
    metadata_writer_check(!is_file($invalid), 'Rejected metadata update left an output file.');

    $thrown = false;
    try {
        Xlsx::updateMetaInfo($source, $invalid, ['unknown_section' => ['value' => 1]]);
    } catch (Throwable) {
        $thrown = true;
    }
    metadata_writer_check($thrown, 'Unknown metadata section was not rejected in strict mode.');

    $thrown = false;
    try {
        Xlsx::updateMetaInfo($source, $invalid, ['document' => ['titel' => 'Typo']]);
    } catch (Throwable) {
        $thrown = true;
    }
    metadata_writer_check($thrown, 'Unknown nested metadata field was not rejected in strict mode.');

    Xlsx::encryptFile($updated, $encrypted, $password, ['mode' => 'agile', 'spin_count' => 1000]);
    $locked = Xlsx::metaInfo($encrypted, ['profile' => 'quick']);
    metadata_writer_check($locked['status'] === 'password_required', 'Encrypted metadata did not require a password.');
    metadata_writer_check($locked['file']['container'] === 'ole-encrypted-ooxml', 'Encrypted source container was not identified.');

    $thrown = false;
    try {
        Xlsx::updateMetaInfo($encrypted, $wrongPassword, ['document' => ['title' => 'Should Fail']], ['password' => 'wrong']);
    } catch (Throwable) {
        $thrown = true;
    }
    metadata_writer_check($thrown, 'Wrong encrypted-workbook password was not rejected.');
    metadata_writer_check(!is_file($wrongPassword), 'Wrong-password update left an output file.');

    Xlsx::updateMetaInfo($encrypted, $encryptedUpdated, [
        'document' => ['title' => 'Encrypted Update'],
    ], ['password' => $password]);
    metadata_writer_check(Xlsx::isEncrypted($encryptedUpdated), 'Encrypted update produced an unencrypted workbook.');
    metadata_writer_check(Xlsx::encryptionMode($encryptedUpdated) === 'agile', 'Encrypted update changed encryption mode.');
    $encryptedInfo = Xlsx::metaInfo($encryptedUpdated, ['profile' => 'quick', 'password' => $password]);
    metadata_writer_check($encryptedInfo['document']['title'] === 'Encrypted Update', 'Encrypted metadata update failed.');
    metadata_writer_check($encryptedInfo['file']['container'] === 'ole-encrypted-ooxml', 'Opened encrypted source container was reported as a plain ZIP.');
    metadata_writer_check($encryptedInfo['file']['inner_package_container'] === 'zip', 'Encrypted OOXML inner package container was not reported.');

    Xlsx::encryptFile($updated, $standardEncrypted, $password, ['mode' => 'standard']);
    Xlsx::updateMetaInfo($standardEncrypted, $standardUpdated, [
        'document' => ['title' => 'Standard Encrypted Update'],
    ], ['xlsx_password' => $password]);
    metadata_writer_check(Xlsx::encryptionMode($standardUpdated) === 'standard', 'Standard encryption mode was not preserved.');
    metadata_writer_check(
        Xlsx::metaInfo($standardUpdated, ['profile' => 'quick', 'xlsx_password' => $password])['document']['title'] === 'Standard Encrypted Update',
        'Standard-encrypted metadata update failed.'
    );

    echo "metadata_writer_smoke passed\n";
} finally {
    foreach ([$source, $updated, $clean, $encrypted, $encryptedUpdated, $invalid, $wrongPassword, $standardEncrypted, $standardUpdated] as $file) {
        @unlink($file);
    }
    @rmdir($directory);
}
