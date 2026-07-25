<?php

declare(strict_types=1);

namespace Mnb\PHPExcel\Reader\SharedStrings;

final class InMemorySharedStringProvider implements SharedStringProviderInterface
{
    /** @param list<string> $strings */
    public function __construct(private array $strings = [])
    {
    }

    public function get(int $index): string
    {
        return $this->strings[$index] ?? '';
    }

    public function count(): int
    {
        return count($this->strings);
    }

    public function mode(): string
    {
        return 'memory';
    }

    public function close(): void
    {
        $this->strings = [];
    }
}
