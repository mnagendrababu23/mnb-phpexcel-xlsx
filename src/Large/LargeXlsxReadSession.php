<?php

declare(strict_types=1);

namespace Mnb\PHPExcel\Large;

use Mnb\PHPExcel\Application\RowTransformerPipeline;

/**
 * Fluent large-XLSX read session. It always streams; it never returns the full workbook.
 */
final class LargeXlsxReadSession
{
    /** @var array<string,mixed> */
    private array $options;

    public function __construct(private string $path, private ?LargeXlsxStreamingReader $reader = null, array $options = [])
    {
        $this->reader = $reader ?? new LargeXlsxStreamingReader();
        $this->options = $options;
    }

    public function sheet(int|string $sheet): self
    {
        $clone = clone $this;
        $clone->options['sheet'] = $sheet;
        return $clone;
    }

    public function withHeader(bool $enabled = true): self
    {
        $clone = clone $this;
        $clone->options['with_header'] = $enabled;
        return $clone;
    }

    public function skipRows(int $rows): self
    {
        $clone = clone $this;
        $clone->options['skip_rows'] = max(0, $rows);
        return $clone;
    }

    public function limitRows(int $rows): self
    {
        $clone = clone $this;
        $clone->options['limit_rows'] = max(0, $rows);
        return $clone;
    }

    /** @param list<int|string> $columns */
    public function onlyColumns(array $columns): self
    {
        $clone = clone $this;
        $clone->options['only_columns'] = $columns;
        return $clone;
    }

    public function timeBudgetSeconds(int $seconds): self
    {
        $clone = clone $this;
        $clone->options['time_budget_seconds'] = max(0, $seconds);
        return $clone;
    }

    public function memoryGuardRatio(float $ratio): self
    {
        $clone = clone $this;
        $clone->options['memory_guard_ratio'] = max(0.1, min(0.95, $ratio));
        return $clone;
    }

    public function maxSharedStringsInMemory(int $count): self
    {
        $clone = clone $this;
        $clone->options['max_shared_strings_in_memory'] = max(0, $count);
        return $clone;
    }


    public function preserveNumericStrings(bool $enabled = true): self
    {
        $clone = clone $this;
        $clone->options['preserve_numeric_strings'] = $enabled;
        return $clone;
    }

    public function convertDates(bool $enabled = true, ?string $format = null): self
    {
        $clone = clone $this;
        $clone->options['convert_dates'] = $enabled;
        if ($format !== null) {
            $clone->options['date_output_format'] = $format;
        }
        return $clone;
    }


    /** @param callable|string $transformer */
    public function transform(callable|string $transformer): self
    {
        $clone = clone $this;
        $clone->options['transformers'][] = $transformer;
        return $clone;
    }

    /** @param list<callable|string> $transformers */
    public function transformers(array $transformers): self
    {
        $clone = clone $this;
        $clone->options['transformers'] = array_values($transformers);
        return $clone;
    }

    public function progress(callable $callback): self
    {
        $clone = clone $this;
        $clone->options['progress'] = $callback;
        return $clone;
    }

    /**
     * @param callable(array<int,array<int|string,mixed>>, array<string,mixed>): (bool|void) $callback
     * @return array<string,mixed>
     */
    public function chunk(int $chunkSize, callable $callback): array
    {
        $options = $this->options;
        $transformers = is_array($options['transformers'] ?? null) ? $options['transformers'] : [];
        if ($transformers !== []) {
            $callback = static function (array $rows, array $state) use ($callback, $transformers): bool|null {
                $rows = RowTransformerPipeline::applyRows($rows, $transformers, $state);
                $result = $callback($rows, $state);
                return is_bool($result) ? $result : null;
            };
        }
        return $this->reader->chunk($this->path, $chunkSize, $callback, $options);
    }

    /**
     * @param callable(array<int|string,mixed>, array<string,mixed>): (bool|void) $callback
     * @return array<string,mixed>
     */
    public function eachRow(callable $callback): array
    {
        $options = $this->options;
        $transformers = is_array($options['transformers'] ?? null) ? $options['transformers'] : [];
        if ($transformers !== []) {
            $callback = static function (array $row, array $state) use ($callback, $transformers): bool|null {
                $row = RowTransformerPipeline::apply($row, $transformers, $state);
                $result = $callback($row, $state);
                return is_bool($result) ? $result : null;
            };
        }
        return $this->reader->eachRow($this->path, $callback, $options);
    }

    /** Alias for chunk(), useful in import controllers/jobs. */
    public function import(int $chunkSize, callable $callback): array
    {
        return $this->chunk($chunkSize, $callback);
    }

    /** @return array<string,mixed> */
    public function options(): array
    {
        return $this->options;
    }
}
