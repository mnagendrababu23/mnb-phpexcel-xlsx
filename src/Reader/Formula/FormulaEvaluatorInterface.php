<?php

declare(strict_types=1);

namespace Mnb\PHPExcel\Reader\Formula;

interface FormulaEvaluatorInterface
{
    public function calculateCell(string $path, int|string $sheet, string $cell): mixed;

    /** @return array<string,mixed> keyed by uppercase cell reference */
    public function calculateRange(string $path, int|string $sheet, string $range): array;
}
