<?php

declare(strict_types=1);

namespace Mnb\PHPExcel\Reader\Formula;

use DateTimeImmutable;
use Mnb\PHPExcel\Reader\XlsxReader;
use Mnb\PHPExcel\Support\Coordinate;
use Mnb\PHPExcel\Support\MnbExcelException;

/**
 * Lightweight native formula evaluator for common business formulas.
 * Unsupported functions can fall back to another evaluator when supplied.
 */
final class NativeFormulaEvaluator implements FormulaEvaluatorInterface
{
    /** @var list<array{type:string,value:mixed}> */
    private array $tokens = [];
    private int $position = 0;
    private string $path = '';
    private int|string $sheet = 1;
    /** @var array<string,bool> */
    private array $stack = [];

    public function __construct(private readonly ?FormulaEvaluatorInterface $fallback = null)
    {
    }

    public function calculateCell(string $path, int|string $sheet, string $cell): mixed
    {
        $this->path = $path;
        $this->sheet = $sheet;
        return $this->resolveCell($sheet, strtoupper(str_replace('$', '', trim($cell))));
    }

    /** @return array<string,mixed> */
    public function calculateRange(string $path, int|string $sheet, string $range): array
    {
        $this->path = $path;
        $this->sheet = $sheet;
        [$start, $end] = $this->normalizeRange($range);
        [$startColumn, $startRow] = Coordinate::splitCellRef($start);
        [$endColumn, $endRow] = Coordinate::splitCellRef($end);
        $values = [];
        for ($row = min($startRow, $endRow); $row <= max($startRow, $endRow); $row++) {
            for ($column = min($startColumn, $endColumn); $column <= max($startColumn, $endColumn); $column++) {
                $cell = Coordinate::columnIndexToName($column) . $row;
                $values[$cell] = $this->resolveCell($sheet, $cell);
            }
        }
        return $values;
    }

    /** Evaluate a formula with a custom cell resolver, useful outside XLSX files and for tests. */
    public function evaluate(string $formula, callable $resolver): mixed
    {
        $previousPath = $this->path;
        $previousSheet = $this->sheet;
        $this->path = '';
        $this->sheet = 1;
        $this->customResolver = $resolver;
        try {
            return $this->evaluateExpression($formula);
        } finally {
            $this->customResolver = null;
            $this->path = $previousPath;
            $this->sheet = $previousSheet;
        }
    }

    /** @var callable(int|string,string):mixed|null */
    private mixed $customResolver = null;

    private function resolveCell(int|string $sheet, string $cell): mixed
    {
        $key = (string) $sheet . '!' . strtoupper($cell);
        if (isset($this->stack[$key])) {
            throw new MnbExcelException('Circular formula reference detected at ' . $key . '.');
        }
        if (is_callable($this->customResolver)) {
            return ($this->customResolver)($sheet, strtoupper($cell));
        }

        $this->stack[$key] = true;
        try {
            $snapshot = (new XlsxReader())->readCellDetails($this->path, $cell, $sheet, ['calculate' => false]);
            if ($snapshot->formula === null || $snapshot->formula === '') {
                return $snapshot->value;
            }
            try {
                $currentSheet = $this->sheet;
                $this->sheet = $sheet;
                $value = $this->evaluateExpression($snapshot->formula);
                $this->sheet = $currentSheet;
                return $value;
            } catch (MnbExcelException $exception) {
                if ($this->fallback !== null) {
                    return $this->fallback->calculateCell($this->path, $sheet, $cell);
                }
                throw $exception;
            }
        } finally {
            unset($this->stack[$key]);
        }
    }

    private function evaluateExpression(string $formula): mixed
    {
        $formula = ltrim(trim($formula), '=');
        $this->tokens = $this->tokenize($formula);
        $this->position = 0;
        $value = $this->parseComparison();
        if (!$this->is('EOF')) {
            throw new MnbExcelException('Unsupported or invalid formula near token: ' . (string) $this->current()['value']);
        }
        return $value;
    }

    /** @return list<array{type:string,value:mixed}> */
    private function tokenize(string $formula): array
    {
        $tokens = [];
        $length = strlen($formula);
        $offset = 0;
        while ($offset < $length) {
            if (preg_match('/\G\s+/A', $formula, $match, 0, $offset) === 1) {
                $offset += strlen($match[0]);
                continue;
            }
            if (preg_match('/\G"((?:[^"]|"")*)"/A', $formula, $match, 0, $offset) === 1) {
                $tokens[] = ['type' => 'STRING', 'value' => str_replace('""', '"', $match[1])];
                $offset += strlen($match[0]);
                continue;
            }
            if (preg_match('/\G(?:\d+(?:\.\d*)?|\.\d+)(?:[Ee][+-]?\d+)?/A', $formula, $match, 0, $offset) === 1) {
                $tokens[] = ['type' => 'NUMBER', 'value' => (float) $match[0]];
                $offset += strlen($match[0]);
                continue;
            }
            if (preg_match('/\G(?:(?:\'((?:[^\']|\'\')+)\'|([A-Za-z_][A-Za-z0-9_.]*))!)?\$?([A-Za-z]{1,3})\$?([1-9][0-9]*)/A', $formula, $match, 0, $offset) === 1) {
                $sheet = $match[1] !== '' ? str_replace("''", "'", $match[1]) : ($match[2] !== '' ? $match[2] : null);
                $tokens[] = ['type' => 'CELL', 'value' => ['sheet' => $sheet, 'cell' => strtoupper($match[3]) . $match[4]]];
                $offset += strlen($match[0]);
                continue;
            }
            if (preg_match('/\G[A-Za-z_][A-Za-z0-9_.]*/A', $formula, $match, 0, $offset) === 1) {
                $tokens[] = ['type' => 'IDENT', 'value' => strtoupper($match[0])];
                $offset += strlen($match[0]);
                continue;
            }
            $two = substr($formula, $offset, 2);
            if (in_array($two, ['<=', '>=', '<>'], true)) {
                $tokens[] = ['type' => 'OP', 'value' => $two];
                $offset += 2;
                continue;
            }
            $char = $formula[$offset];
            $type = match ($char) {
                '(' => 'LPAREN', ')' => 'RPAREN', ',' => 'COMMA', ':' => 'COLON',
                '+', '-', '*', '/', '^', '&', '=', '<', '>', '%' => 'OP',
                default => null,
            };
            if ($type === null) {
                throw new MnbExcelException('Unsupported formula character: ' . $char);
            }
            $tokens[] = ['type' => $type, 'value' => $char];
            $offset++;
        }
        $tokens[] = ['type' => 'EOF', 'value' => null];
        return $tokens;
    }

    private function parseComparison(): mixed
    {
        $left = $this->parseConcat();
        while ($this->is('OP') && in_array($this->current()['value'], ['=', '<>', '<', '>', '<=', '>='], true)) {
            $operator = (string) $this->consume()['value'];
            $right = $this->parseConcat();
            $left = match ($operator) {
                '=' => $left == $right,
                '<>' => $left != $right,
                '<' => $left < $right,
                '>' => $left > $right,
                '<=' => $left <= $right,
                '>=' => $left >= $right,
            };
        }
        return $left;
    }

    private function parseConcat(): mixed
    {
        $left = $this->parseAdditive();
        while ($this->is('OP', '&')) {
            $this->consume();
            $left = $this->toText($left) . $this->toText($this->parseAdditive());
        }
        return $left;
    }

    private function parseAdditive(): mixed
    {
        $left = $this->parseMultiplicative();
        while ($this->is('OP') && in_array($this->current()['value'], ['+', '-'], true)) {
            $operator = (string) $this->consume()['value'];
            $right = $this->parseMultiplicative();
            $left = $operator === '+' ? $this->number($left) + $this->number($right) : $this->number($left) - $this->number($right);
        }
        return $left;
    }

    private function parseMultiplicative(): mixed
    {
        $left = $this->parsePower();
        while ($this->is('OP') && in_array($this->current()['value'], ['*', '/'], true)) {
            $operator = (string) $this->consume()['value'];
            $right = $this->number($this->parsePower());
            if ($operator === '/' && $right == 0.0) {
                throw new MnbExcelException('Formula division by zero.');
            }
            $left = $operator === '*' ? $this->number($left) * $right : $this->number($left) / $right;
        }
        return $left;
    }

    private function parsePower(): mixed
    {
        $left = $this->parseUnary();
        if ($this->is('OP', '^')) {
            $this->consume();
            $left = $this->number($left) ** $this->number($this->parsePower());
        }
        return $left;
    }

    private function parseUnary(): mixed
    {
        if ($this->is('OP', '+')) {
            $this->consume();
            return $this->number($this->parseUnary());
        }
        if ($this->is('OP', '-')) {
            $this->consume();
            return -$this->number($this->parseUnary());
        }
        $value = $this->parsePrimary();
        if ($this->is('OP', '%')) {
            $this->consume();
            return $this->number($value) / 100;
        }
        return $value;
    }

    private function parsePrimary(): mixed
    {
        if ($this->is('NUMBER') || $this->is('STRING')) {
            return $this->consume()['value'];
        }
        if ($this->is('IDENT')) {
            $name = (string) $this->consume()['value'];
            if ($name === 'TRUE') {
                return true;
            }
            if ($name === 'FALSE') {
                return false;
            }
            if ($this->is('LPAREN')) {
                $this->consume();
                $args = [];
                if (!$this->is('RPAREN')) {
                    do {
                        $args[] = $this->parseComparison();
                        if (!$this->is('COMMA')) {
                            break;
                        }
                        $this->consume();
                    } while (true);
                }
                $this->expect('RPAREN');
                return $this->callFunction($name, $args);
            }
            throw new MnbExcelException('Unsupported formula name: ' . $name);
        }
        if ($this->is('CELL')) {
            $start = $this->consume()['value'];
            if ($this->is('COLON')) {
                $this->consume();
                $end = $this->expect('CELL')['value'];
                return $this->resolveRange($start, $end);
            }
            return $this->resolveCell($start['sheet'] ?? $this->sheet, (string) $start['cell']);
        }
        if ($this->is('LPAREN')) {
            $this->consume();
            $value = $this->parseComparison();
            $this->expect('RPAREN');
            return $value;
        }
        throw new MnbExcelException('Unexpected formula token: ' . (string) $this->current()['type']);
    }

    /** @param array{sheet:mixed,cell:string} $start @param array{sheet:mixed,cell:string} $end @return list<mixed> */
    private function resolveRange(array $start, array $end): array
    {
        $sheet = $start['sheet'] ?? $end['sheet'] ?? $this->sheet;
        [$startColumn, $startRow] = Coordinate::splitCellRef((string) $start['cell']);
        [$endColumn, $endRow] = Coordinate::splitCellRef((string) $end['cell']);
        $values = [];
        for ($row = min($startRow, $endRow); $row <= max($startRow, $endRow); $row++) {
            for ($column = min($startColumn, $endColumn); $column <= max($startColumn, $endColumn); $column++) {
                $values[] = $this->resolveCell($sheet, Coordinate::columnIndexToName($column) . $row);
            }
        }
        return $values;
    }

    /** @param list<mixed> $args */
    private function callFunction(string $name, array $args): mixed
    {
        $flat = $this->flatten($args);
        $numbers = array_values(array_map([$this, 'number'], array_filter($flat, static fn (mixed $value): bool => is_numeric($value) || is_bool($value))));
        return match ($name) {
            'SUM' => array_sum($numbers),
            'AVERAGE' => $numbers === [] ? 0.0 : array_sum($numbers) / count($numbers),
            'MIN' => $numbers === [] ? 0.0 : min($numbers),
            'MAX' => $numbers === [] ? 0.0 : max($numbers),
            'COUNT' => count($numbers),
            'COUNTA' => count(array_filter($flat, static fn (mixed $v): bool => $v !== null && $v !== '')),
            'IF' => $this->truthy($args[0] ?? false) ? ($args[1] ?? true) : ($args[2] ?? false),
            'IFERROR' => $args[0] ?? ($args[1] ?? null),
            'AND' => !in_array(false, array_map([$this, 'truthy'], $flat), true),
            'OR' => in_array(true, array_map([$this, 'truthy'], $flat), true),
            'NOT' => !$this->truthy($args[0] ?? false),
            'ABS' => abs($this->number($args[0] ?? 0)),
            'SQRT' => sqrt($this->number($args[0] ?? 0)),
            'POWER' => $this->number($args[0] ?? 0) ** $this->number($args[1] ?? 0),
            'ROUND' => round($this->number($args[0] ?? 0), (int) $this->number($args[1] ?? 0)),
            'ROUNDUP' => $this->roundDirectional($this->number($args[0] ?? 0), (int) $this->number($args[1] ?? 0), true),
            'ROUNDDOWN' => $this->roundDirectional($this->number($args[0] ?? 0), (int) $this->number($args[1] ?? 0), false),
            'LEN' => $this->stringLength($this->toText($args[0] ?? '')),
            'LOWER' => $this->lower($this->toText($args[0] ?? '')),
            'UPPER' => $this->upper($this->toText($args[0] ?? '')),
            'TRIM' => preg_replace('/\s+/u', ' ', trim($this->toText($args[0] ?? ''))) ?? '',
            'LEFT' => $this->substring($this->toText($args[0] ?? ''), 0, max(0, (int) $this->number($args[1] ?? 1))),
            'RIGHT' => $this->right($this->toText($args[0] ?? ''), max(0, (int) $this->number($args[1] ?? 1))),
            'MID' => $this->substring($this->toText($args[0] ?? ''), max(0, (int) $this->number($args[1] ?? 1) - 1), max(0, (int) $this->number($args[2] ?? 0))),
            'CONCAT', 'CONCATENATE' => implode('', array_map([$this, 'toText'], $flat)),
            'DATE' => sprintf('%04d-%02d-%02d', (int) $this->number($args[0] ?? 0), (int) $this->number($args[1] ?? 1), (int) $this->number($args[2] ?? 1)),
            'TODAY' => (new DateTimeImmutable('today'))->format('Y-m-d'),
            'NOW' => (new DateTimeImmutable('now'))->format('Y-m-d H:i:s'),
            'YEAR' => (int) (new DateTimeImmutable($this->toText($args[0] ?? 'now')))->format('Y'),
            'MONTH' => (int) (new DateTimeImmutable($this->toText($args[0] ?? 'now')))->format('n'),
            'DAY' => (int) (new DateTimeImmutable($this->toText($args[0] ?? 'now')))->format('j'),
            default => throw new MnbExcelException('Native formula engine does not support function ' . $name . '.'),
        };
    }

    /** @param list<mixed> $values @return list<mixed> */
    private function flatten(array $values): array
    {
        $result = [];
        array_walk_recursive($values, static function (mixed $value) use (&$result): void { $result[] = $value; });
        return $result;
    }

    private function number(mixed $value): float
    {
        if (is_bool($value)) {
            return $value ? 1.0 : 0.0;
        }
        if ($value === null || $value === '') {
            return 0.0;
        }
        if (!is_numeric($value)) {
            throw new MnbExcelException('Formula expected a numeric value, received: ' . $this->toText($value));
        }
        return (float) $value;
    }

    private function truthy(mixed $value): bool
    {
        if (is_array($value)) {
            return $value !== [] && $this->truthy(reset($value));
        }
        return !($value === false || $value === null || $value === '' || $value === 0 || $value === 0.0 || $value === '0');
    }

    private function toText(mixed $value): string
    {
        if (is_array($value)) {
            return implode('', array_map([$this, 'toText'], $this->flatten($value)));
        }
        if (is_bool($value)) {
            return $value ? 'TRUE' : 'FALSE';
        }
        return $value === null ? '' : (string) $value;
    }

    private function roundDirectional(float $value, int $digits, bool $away): float
    {
        $factor = 10 ** $digits;
        $scaled = $value * $factor;
        $rounded = $away ? ($scaled >= 0 ? ceil($scaled) : floor($scaled)) : ($scaled >= 0 ? floor($scaled) : ceil($scaled));
        return $rounded / $factor;
    }

    private function stringLength(string $value): int
    {
        return function_exists('mb_strlen') ? mb_strlen($value, 'UTF-8') : strlen($value);
    }

    private function lower(string $value): string
    {
        return function_exists('mb_strtolower') ? mb_strtolower($value, 'UTF-8') : strtolower($value);
    }

    private function upper(string $value): string
    {
        return function_exists('mb_strtoupper') ? mb_strtoupper($value, 'UTF-8') : strtoupper($value);
    }

    private function substring(string $value, int $start, int $length): string
    {
        return function_exists('mb_substr') ? mb_substr($value, $start, $length, 'UTF-8') : substr($value, $start, $length);
    }

    private function right(string $value, int $length): string
    {
        $total = $this->stringLength($value);
        return $this->substring($value, max(0, $total - $length), $length);
    }

    /** @return array{string,string} */
    private function normalizeRange(string $range): array
    {
        $range = strtoupper(str_replace('$', '', trim($range)));
        if (preg_match('/^([A-Z]+\d+)(?::([A-Z]+\d+))?$/', $range, $match) !== 1) {
            throw new MnbExcelException('Invalid cell range: ' . $range);
        }
        return [$match[1], $match[2] ?? $match[1]];
    }

    /** @return array{type:string,value:mixed} */
    private function current(): array
    {
        return $this->tokens[$this->position] ?? ['type' => 'EOF', 'value' => null];
    }

    private function is(string $type, mixed $value = null): bool
    {
        $token = $this->current();
        return $token['type'] === $type && ($value === null || $token['value'] === $value);
    }

    /** @return array{type:string,value:mixed} */
    private function consume(): array
    {
        return $this->tokens[$this->position++] ?? ['type' => 'EOF', 'value' => null];
    }

    /** @return array{type:string,value:mixed} */
    private function expect(string $type): array
    {
        if (!$this->is($type)) {
            throw new MnbExcelException('Expected formula token ' . $type . ', got ' . $this->current()['type'] . '.');
        }
        return $this->consume();
    }
}
