<?php

declare(strict_types=1);

namespace Mnb\PHPExcel\Reader\SharedStrings;

interface SharedStringProviderInterface
{
    public function get(int $index): string;
    public function count(): int;
    public function mode(): string;
    public function close(): void;
}
