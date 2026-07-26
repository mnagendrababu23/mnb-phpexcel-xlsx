<?php

declare(strict_types=1);

namespace Mnb\PHPExcel\Reader\Formula;

final class FormulaEvaluatorFactory
{
    public static function create(bool $preferNative = false): FormulaEvaluatorInterface
    {
        $phpSpreadsheet = new PhpSpreadsheetFormulaEvaluator();
        if (!$preferNative && $phpSpreadsheet->available()) {
            return $phpSpreadsheet;
        }
        return new NativeFormulaEvaluator($phpSpreadsheet->available() ? $phpSpreadsheet : null);
    }
}
