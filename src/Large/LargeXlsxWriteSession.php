<?php

declare(strict_types=1);

namespace Mnb\PHPExcel\Large;

/**
 * Fluent session for large XLSX/CSV-ZIP exports. The source iterable is streamed
 * exactly once, so call either save() or saveCsvZip() for a session.
 */
final class LargeXlsxWriteSession
{
    /** @var array<string,mixed> */
    private array $options;

    /** @param iterable<array<int|string,mixed>> $rows */
    public function __construct(private iterable $rows, private ?LargeXlsxStreamingWriter $writer = null, array $options = [])
    {
        $this->writer = $writer ?? new LargeXlsxStreamingWriter();
        $this->options = $options;
    }

    public function sheetName(string $name): self
    {
        $clone = clone $this;
        $clone->options['sheet_name'] = $name;
        return $clone;
    }

    /** @param list<string>|bool $header */
    public function withHeader(array|bool $header = true): self
    {
        $clone = clone $this;
        $clone->options['with_header'] = $header !== false;
        if (is_array($header)) {
            $clone->options['headers'] = array_values($header);
        }
        return $clone;
    }

    /** @param list<string> $headers */
    public function headers(array $headers): self
    {
        return $this->withHeader($headers);
    }

    public function autoSplitSheets(bool $enabled = true, int $maxRowsPerSheet = 1048576): self
    {
        $clone = clone $this;
        $clone->options['auto_split_sheets'] = $enabled;
        $clone->options['max_rows_per_sheet'] = $maxRowsPerSheet;
        return $clone;
    }

    public function maxRowsPerSheet(int $rows): self
    {
        $clone = clone $this;
        $clone->options['max_rows_per_sheet'] = max(1, $rows);
        return $clone;
    }

    public function progress(callable $callback, int $everyRows = 1000): self
    {
        $clone = clone $this;
        $clone->options['progress'] = $callback;
        $clone->options['progress_every'] = max(1, $everyRows);
        return $clone;
    }

    public function freezeHeader(bool $enabled = true): self
    {
        $clone = clone $this;
        $clone->options['freeze_header'] = $enabled;
        return $clone;
    }

    public function autoFilter(bool $enabled = true): self
    {
        $clone = clone $this;
        $clone->options['auto_filter'] = $enabled;
        return $clone;
    }

    public function validateIntegrity(bool $enabled = true): self
    {
        $clone = clone $this;
        $clone->options['validate_integrity'] = $enabled;
        return $clone;
    }

    public function formulaPolicy(string $policy): self
    {
        $clone = clone $this;
        $clone->options['formula_policy'] = $policy;
        return $clone;
    }

    public function textPolicy(string $policy): self
    {
        $clone = clone $this;
        $clone->options['formula_text_policy'] = $policy;
        $clone->options['csv_injection_policy'] = $policy;
        return $clone;
    }

    /** @param array<int|string,string> $formats */
    public function columnFormats(array $formats): self
    {
        $clone = clone $this;
        $clone->options['column_formats'] = $formats;
        return $clone;
    }

    public function formatColumn(int|string $column, string $format): self
    {
        $clone = clone $this;
        $formats = is_array($clone->options['column_formats'] ?? null) ? $clone->options['column_formats'] : [];
        $formats[$column] = $format;
        $clone->options['column_formats'] = $formats;
        return $clone;
    }

    public function csvRowsPerFile(int $rows): self
    {
        $clone = clone $this;
        $clone->options['rows_per_file'] = max(1, $rows);
        return $clone;
    }

    /** @return array<string,mixed> */
    public function save(string $path): array
    {
        return $this->writer->write($this->rows, $path, $this->options);
    }

    /** @return array<string,mixed> */
    public function saveCsvZip(string $path): array
    {
        return $this->writer->saveCsvZip($this->rows, $path, $this->options);
    }

    /** @return array<string,mixed> */
    public function options(): array
    {
        return $this->options;
    }
}
