<?php

declare(strict_types=1);

namespace Mnb\PHPExcel\Reader\Formula;

final class FormulaEvaluatorFactory
{
    public static function create(bool $preferNative = true): FormulaEvaluatorInterface
    {
        return new NativeFormulaEvaluator();
    }
}
