<?php

declare(strict_types=1);

$repositories = dirname(__DIR__, 2);
$sources = [
    dirname(__DIR__) . '/src/',
    $repositories . '/mnb-phpexcel-core/src/',
];
spl_autoload_register(static function (string $class) use ($sources): void {
    $prefix = 'Mnb\\PHPExcel\\';
    if (!str_starts_with($class, $prefix)) {
        return;
    }
    $relative = str_replace('\\', '/', substr($class, strlen($prefix))) . '.php';
    foreach ($sources as $source) {
        $file = $source . $relative;
        if (is_file($file)) {
            require $file;
            return;
        }
    }
});

use Mnb\PHPExcel\Format\Xlsx;
use Mnb\PHPExcel\Reader\XlsxReader;
use Mnb\PHPExcel\Support\ErrorCode;
use Mnb\PHPExcel\Support\SheetSelectionException;

$path = tempnam(sys_get_temp_dir(), 'mnb-xlsx-sheet-');
if ($path === false) {
    throw new RuntimeException('Unable to create XLSX test path.');
}
@unlink($path);
$path .= '.xlsx';

Xlsx::write([
    ['parameter' => 'alpha', 'value' => 1],
    ['parameter' => 'beta', 'value' => 2],
], $path, ['sheet_name' => 'ALL_PARAMETERS', 'with_header' => true]);

$expect = static function (callable $callback, string $code, string $messagePart): void {
    try {
        $callback();
    } catch (SheetSelectionException $e) {
        assert($e->getErrorCode() === $code, $e->getErrorCode());
        assert(str_contains($e->getMessage(), $messagePart), $e->getMessage());
        assert(str_contains($e->getMessage(), basename(__FILE__)), $e->getMessage());
        return;
    }

    throw new RuntimeException('Expected SheetSelectionException was not thrown.');
};

$expect(static fn() => Xlsx::read($path)->sheet(), ErrorCode::SHEET_SELECTION_REQUIRED, 'omit ->sheet() entirely');
$expect(static fn() => Xlsx::read($path)->sheet(0), ErrorCode::SHEET_INDEX_INVALID, '1-based');
$expect(static fn() => Xlsx::read($path)->sheet('Sheet1'), ErrorCode::SHEET_NOT_FOUND, 'ALL_PARAMETERS');
$expect(static fn() => (new XlsxReader())->readSheet($path, 0), ErrorCode::SHEET_INDEX_INVALID, '1-based');
$expect(static fn() => (new XlsxReader())->readSheet($path, 'Sheet1'), ErrorCode::SHEET_NOT_FOUND, 'ALL_PARAMETERS');

$rows = Xlsx::read($path)->sheet('all_parameters')->withHeaderRow(1)->toArray();
assert(count($rows) === 2);
assert($rows[0]['parameter'] === 'alpha');

@unlink($path);
echo "xlsx sheet selection errors: ok\n";
