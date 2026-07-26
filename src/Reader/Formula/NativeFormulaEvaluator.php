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

    public function __construct(
        private readonly ?FormulaEvaluatorInterface $fallback = null,
        private readonly ?FormulaFunctionRegistry $functions = null
    ) {
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

    /** @param array{sheet:mixed,cell:string} $start @param array{sheet:mixed,cell:string} $end @return list<list<mixed>> */
    private function resolveRange(array $start, array $end): array
    {
        $sheet = $start['sheet'] ?? $end['sheet'] ?? $this->sheet;
        [$startColumn, $startRow] = Coordinate::splitCellRef((string) $start['cell']);
        [$endColumn, $endRow] = Coordinate::splitCellRef((string) $end['cell']);
        $values = [];
        for ($row = min($startRow, $endRow); $row <= max($startRow, $endRow); $row++) {
            $line = [];
            for ($column = min($startColumn, $endColumn); $column <= max($startColumn, $endColumn); $column++) {
                $line[] = $this->resolveCell($sheet, Coordinate::columnIndexToName($column) . $row);
            }
            $values[] = $line;
        }
        return $values;
    }

    /** @param list<mixed> $args */
    private function callFunction(string $name, array $args): mixed
    {
        if ($this->functions?->has($name)) {
            return $this->functions->call($name, $args);
        }
        $flat = $this->flatten($args);
        $numbers = array_values(array_map([$this, 'number'], array_filter($flat, static fn (mixed $value): bool => is_numeric($value) || is_bool($value))));
        return match ($name) {
            'SUM' => array_sum($numbers),
            'AVERAGE' => $numbers === [] ? 0.0 : array_sum($numbers) / count($numbers),
            'MIN' => $numbers === [] ? 0.0 : min($numbers),
            'MAX' => $numbers === [] ? 0.0 : max($numbers),
            'COUNT' => count($numbers),
            'COUNTA' => count(array_filter($flat, static fn (mixed $v): bool => $v !== null && $v !== '')),
            'COUNTBLANK' => count(array_filter($flat, static fn (mixed $v): bool => $v === null || $v === '')),
            'SUMIF' => $this->conditionalAggregate($args, 'sum'),
            'AVERAGEIF' => $this->conditionalAggregate($args, 'average'),
            'COUNTIF' => $this->conditionalAggregate($args, 'count'),
            'SUMIFS' => $this->multiConditionalAggregate($args, 'sum'),
            'AVERAGEIFS' => $this->multiConditionalAggregate($args, 'average'),
            'COUNTIFS' => $this->multiConditionalAggregate($args, 'count'),
            'MINIFS' => $this->multiConditionalAggregate($args, 'min'),
            'MAXIFS' => $this->multiConditionalAggregate($args, 'max'),
            'IF' => $this->truthy($args[0] ?? false) ? ($args[1] ?? true) : ($args[2] ?? false),
            'IFS' => $this->ifs($args),
            'IFERROR', 'IFNA' => $args[0] ?? ($args[1] ?? null),
            'AND' => !in_array(false, array_map([$this, 'truthy'], $flat), true),
            'OR' => in_array(true, array_map([$this, 'truthy'], $flat), true),
            'XOR' => count(array_filter(array_map([$this, 'truthy'], $flat))) % 2 === 1,
            'NOT' => !$this->truthy($args[0] ?? false),
            'ISBLANK' => ($args[0] ?? null) === null || ($args[0] ?? null) === '',
            'ISNUMBER' => is_int($args[0] ?? null) || is_float($args[0] ?? null) || (is_string($args[0] ?? null) && is_numeric($args[0])),
            'ISTEXT' => is_string($args[0] ?? null) && !is_numeric($args[0]),
            'ISLOGICAL' => is_bool($args[0] ?? null),
            'ISERROR', 'ISERR', 'ISNA' => false,
            'ABS' => abs($this->number($args[0] ?? 0)),
            'SQRT' => sqrt($this->number($args[0] ?? 0)),
            'POWER' => $this->number($args[0] ?? 0) ** $this->number($args[1] ?? 0),
            'EXP' => exp($this->number($args[0] ?? 0)),
            'LN' => log($this->number($args[0] ?? 0)),
            'LOG' => log($this->number($args[0] ?? 0), $this->number($args[1] ?? 10)),
            'LOG10' => log10($this->number($args[0] ?? 0)),
            'PI' => M_PI,
            'SIN' => sin($this->number($args[0] ?? 0)),
            'COS' => cos($this->number($args[0] ?? 0)),
            'TAN' => tan($this->number($args[0] ?? 0)),
            'ASIN' => asin($this->number($args[0] ?? 0)),
            'ACOS' => acos($this->number($args[0] ?? 0)),
            'ATAN' => atan($this->number($args[0] ?? 0)),
            'DEGREES' => rad2deg($this->number($args[0] ?? 0)),
            'RADIANS' => deg2rad($this->number($args[0] ?? 0)),
            'ROUND' => round($this->number($args[0] ?? 0), (int) $this->number($args[1] ?? 0)),
            'ROUNDUP' => $this->roundDirectional($this->number($args[0] ?? 0), (int) $this->number($args[1] ?? 0), true),
            'ROUNDDOWN' => $this->roundDirectional($this->number($args[0] ?? 0), (int) $this->number($args[1] ?? 0), false),
            'INT' => floor($this->number($args[0] ?? 0)),
            'TRUNC' => $this->truncate($this->number($args[0] ?? 0), (int) $this->number($args[1] ?? 0)),
            'MOD' => fmod($this->number($args[0] ?? 0), $this->number($args[1] ?? 1)),
            'SIGN' => $this->number($args[0] ?? 0) <=> 0.0,
            'CEILING', 'CEILING.MATH' => $this->ceilingFloor($args, true),
            'FLOOR', 'FLOOR.MATH' => $this->ceilingFloor($args, false),
            'MROUND' => $this->multipleRound($this->number($args[0] ?? 0), $this->number($args[1] ?? 1)),
            'RAND' => mt_rand() / mt_getrandmax(),
            'RANDBETWEEN' => random_int((int)$this->number($args[0] ?? 0), (int)$this->number($args[1] ?? 0)),
            'LEN' => $this->stringLength($this->toText($args[0] ?? '')),
            'LOWER' => $this->lower($this->toText($args[0] ?? '')),
            'UPPER' => $this->upper($this->toText($args[0] ?? '')),
            'PROPER' => $this->proper($this->toText($args[0] ?? '')),
            'TRIM' => preg_replace('/\s+/u', ' ', trim($this->toText($args[0] ?? ''))) ?? '',
            'CLEAN' => preg_replace('/[\x00-\x1F\x7F]/u', '', $this->toText($args[0] ?? '')) ?? '',
            'LEFT' => $this->substring($this->toText($args[0] ?? ''), 0, max(0, (int) $this->number($args[1] ?? 1))),
            'RIGHT' => $this->right($this->toText($args[0] ?? ''), max(0, (int) $this->number($args[1] ?? 1))),
            'MID' => $this->substring($this->toText($args[0] ?? ''), max(0, (int) $this->number($args[1] ?? 1) - 1), max(0, (int) $this->number($args[2] ?? 0))),
            'CONCAT', 'CONCATENATE' => implode('', array_map([$this, 'toText'], $flat)),
            'TEXTJOIN' => $this->textJoin($args),
            'SUBSTITUTE' => $this->substitute($args),
            'REPLACE' => $this->replaceText($args),
            'FIND' => $this->findText($args, true),
            'SEARCH' => $this->findText($args, false),
            'EXACT' => $this->toText($args[0] ?? '') === $this->toText($args[1] ?? ''),
            'VALUE' => $this->number(str_replace([',', '$'], '', $this->toText($args[0] ?? 0))),
            'REPT' => str_repeat($this->toText($args[0] ?? ''), max(0, (int)$this->number($args[1] ?? 0))),
            'CHAR' => chr((int)$this->number($args[0] ?? 0)),
            'CODE' => ord($this->toText($args[0] ?? "\0")[0] ?? "\0"),
            'VLOOKUP' => $this->vlookup($args),
            'HLOOKUP' => $this->hlookup($args),
            'XLOOKUP' => $this->xlookup($args),
            'INDEX' => $this->indexValue($args),
            'MATCH', 'XMATCH' => $this->matchValue($args),
            'CHOOSE' => $args[max(1, (int)$this->number($args[0] ?? 1))] ?? null,
            'DATE' => $this->dateValue((int)$this->number($args[0] ?? 0),(int)$this->number($args[1] ?? 1),(int)$this->number($args[2] ?? 1))->format('Y-m-d'),
            'DATEVALUE' => (new DateTimeImmutable($this->toText($args[0] ?? 'now')))->format('Y-m-d'),
            'TIME' => sprintf('%02d:%02d:%02d',(int)$this->number($args[0]??0)%24,(int)$this->number($args[1]??0)%60,(int)$this->number($args[2]??0)%60),
            'TODAY' => (new DateTimeImmutable('today'))->format('Y-m-d'),
            'NOW' => (new DateTimeImmutable('now'))->format('Y-m-d H:i:s'),
            'YEAR' => (int) (new DateTimeImmutable($this->toText($args[0] ?? 'now')))->format('Y'),
            'MONTH' => (int) (new DateTimeImmutable($this->toText($args[0] ?? 'now')))->format('n'),
            'DAY' => (int) (new DateTimeImmutable($this->toText($args[0] ?? 'now')))->format('j'),
            'HOUR' => (int) (new DateTimeImmutable($this->toText($args[0] ?? 'now')))->format('G'),
            'MINUTE' => (int) (new DateTimeImmutable($this->toText($args[0] ?? 'now')))->format('i'),
            'SECOND' => (int) (new DateTimeImmutable($this->toText($args[0] ?? 'now')))->format('s'),
            'EDATE' => $this->shiftMonth($args, false),
            'EOMONTH' => $this->shiftMonth($args, true),
            'DAYS' => $this->dateDiffDays($args[1] ?? 'now', $args[0] ?? 'now'),
            'NETWORKDAYS', 'NETWORKDAYS.INTL' => $this->networkDays($args),
            default => throw new MnbExcelException('Native formula engine does not support function ' . $name . '.'),
        };
    }

    /** @param list<mixed> $args */
    private function conditionalAggregate(array $args,string $mode):mixed
    {
        $criteriaValues=$this->flatten([(array)($args[0]??[])]);$criterion=$args[1]??null;$target=$this->flatten([(array)($args[2]??$args[0]??[])]);$matches=[];
        foreach($criteriaValues as $i=>$value){if($this->matchesCriterion($value,$criterion))$matches[]=$target[$i]??null;}
        return $this->aggregate($matches,$mode);
    }
    /** @param list<mixed> $args */
    private function multiConditionalAggregate(array $args,string $mode):mixed
    {
        if($mode==='count'){$target=$this->flatten([(array)($args[0]??[])]);$offset=0;}else{$target=$this->flatten([(array)($args[0]??[])]);$offset=1;}
        $criteria=[];for($i=$offset;$i<count($args);$i+=2){$criteria[]=[$this->flatten([(array)($args[$i]??[])]),$args[$i+1]??null];}
        $matches=[];foreach($target as $index=>$value){$ok=true;foreach($criteria as [$range,$criterion]){if(!$this->matchesCriterion($range[$index]??null,$criterion)){$ok=false;break;}}if($ok)$matches[]=$value;}
        return $this->aggregate($matches,$mode);
    }
    /** @param list<mixed> $values */
    private function aggregate(array $values,string $mode):mixed{$numbers=array_values(array_map([$this,'number'],array_filter($values,fn($v)=>is_numeric($v)||is_bool($v))));return match($mode){'count'=>count($values),'average'=>$numbers===[]?0.0:array_sum($numbers)/count($numbers),'min'=>$numbers===[]?0.0:min($numbers),'max'=>$numbers===[]?0.0:max($numbers),default=>array_sum($numbers)};}
    private function matchesCriterion(mixed $value,mixed $criterion):bool
    {
        if(!is_string($criterion))return $value==$criterion;$operator='=';$expected=$criterion;if(preg_match('/^(<=|>=|<>|=|<|>)(.*)$/s',$criterion,$m)===1){$operator=$m[1];$expected=$m[2];}
        if(str_contains($expected,'*')||str_contains($expected,'?')){$pattern='/^'.str_replace(['\\*','\\?'],['.*','.'],preg_quote($expected,'/')).'$/iu';$equal=preg_match($pattern,$this->toText($value))===1;return $operator==='<>'?!$equal:$equal;}
        $left=is_numeric($value)&&is_numeric($expected)?(float)$value:$this->toText($value);$right=is_numeric($value)&&is_numeric($expected)?(float)$expected:$expected;return match($operator){'<>' =>$left!=$right,'<' =>$left<$right,'>' =>$left>$right,'<='=>$left<=$right,'>='=>$left>=$right,default=>$left==$right};
    }
    /** @param list<mixed> $args */
    private function ifs(array $args):mixed{for($i=0;$i<count($args)-1;$i+=2)if($this->truthy($args[$i]))return$args[$i+1];throw new MnbExcelException('IFS found no true condition.');}
    /** @param list<mixed> $args */
    private function textJoin(array $args):string{$delimiter=$this->toText($args[0]??'');$ignore=$this->truthy($args[1]??true);$items=$this->flatten(array_slice($args,2));if($ignore)$items=array_values(array_filter($items,fn($v)=>$v!==null&&$v!==''));return implode($delimiter,array_map([$this,'toText'],$items));}
    /** @param list<mixed> $args */
    private function substitute(array $args):string{$text=$this->toText($args[0]??'');$old=$this->toText($args[1]??'');$new=$this->toText($args[2]??'');$instance=(int)$this->number($args[3]??0);if($instance<=0)return str_replace($old,$new,$text);$offset=0;$count=0;while(($pos=strpos($text,$old,$offset))!==false){$count++;if($count===$instance)return substr($text,0,$pos).$new.substr($text,$pos+strlen($old));$offset=$pos+strlen($old);}return$text;}
    /** @param list<mixed> $args */
    private function replaceText(array $args):string{$text=$this->toText($args[0]??'');$start=max(0,(int)$this->number($args[1]??1)-1);$length=max(0,(int)$this->number($args[2]??0));return$this->substring($text,0,$start).$this->toText($args[3]??'').$this->substring($text,$start+$length,max(0,$this->stringLength($text)));}
    /** @param list<mixed> $args */
    private function findText(array $args,bool $caseSensitive):int{$needle=$this->toText($args[0]??'');$haystack=$this->toText($args[1]??'');$start=max(0,(int)$this->number($args[2]??1)-1);$segment=$this->substring($haystack,$start,max(0,$this->stringLength($haystack)));$pos=$caseSensitive?strpos($segment,$needle):stripos($segment,$needle);if($pos===false)throw new MnbExcelException('Formula text was not found.');return$start+$pos+1;}
    /** @param list<mixed> $args */
    private function vlookup(array $args):mixed{$table=$this->matrix($args[1]??[]);$col=max(1,(int)$this->number($args[2]??1));$approx=$this->truthy($args[3]??true);$found=null;foreach($table as $row){if(!array_key_exists(0,$row))continue;if($row[0]==($args[0]??null))return$row[$col-1]??null;if($approx&&$row[0]<=($args[0]??null))$found=$row[$col-1]??null;}if($found!==null)return$found;throw new MnbExcelException('VLOOKUP value not found.');}
    /** @param list<mixed> $args */
    private function hlookup(array $args):mixed{$table=$this->matrix($args[1]??[]);if($table===[])throw new MnbExcelException('HLOOKUP table is empty.');$row=max(1,(int)$this->number($args[2]??1));$index=array_search($args[0]??null,$table[0],false);if($index===false)throw new MnbExcelException('HLOOKUP value not found.');return$table[$row-1][$index]??null;}
    /** @param list<mixed> $args */
    private function xlookup(array $args):mixed{$lookup=$this->flatten([(array)($args[1]??[])]);$returns=$this->flatten([(array)($args[2]??[])]);foreach($lookup as $i=>$value)if($value==($args[0]??null))return$returns[$i]??null;if(array_key_exists(3,$args))return$args[3];throw new MnbExcelException('XLOOKUP value not found.');}
    /** @param list<mixed> $args */
    private function indexValue(array $args):mixed{$matrix=$this->matrix($args[0]??[]);$row=max(1,(int)$this->number($args[1]??1));$col=max(1,(int)$this->number($args[2]??1));return$matrix[$row-1][$col-1]??null;}
    /** @param list<mixed> $args */
    private function matchValue(array $args):int{$values=$this->flatten([(array)($args[1]??[])]);$type=(int)$this->number($args[2]??0);$found=null;foreach($values as $i=>$value){if($value==($args[0]??null))return$i+1;if($type===1&&$value<=($args[0]??null))$found=$i+1;if($type===-1&&$value>=($args[0]??null))$found=$i+1;}if($found!==null)return$found;throw new MnbExcelException('MATCH value not found.');}
    /** @return list<list<mixed>> */
    private function matrix(mixed $value):array{if(!is_array($value))return[[$value]];if($value===[])return[];return is_array(reset($value))?array_map(fn($row)=>array_values((array)$row),array_values($value)):[array_values($value)];}
    private function truncate(float $value,int $digits):float{$factor=10**$digits;return($value>=0?floor($value*$factor):ceil($value*$factor))/$factor;}
    /** @param list<mixed> $args */
    private function ceilingFloor(array $args,bool $up):float{$value=$this->number($args[0]??0);$significance=abs($this->number($args[1]??1));if($significance==0.0)return 0.0;return($up?ceil($value/$significance):floor($value/$significance))*$significance;}
    private function multipleRound(float $value,float $multiple):float{if($multiple==0.0)return 0.0;return round($value/$multiple)*$multiple;}
    private function proper(string $value):string{return function_exists('mb_convert_case')?mb_convert_case($value,MB_CASE_TITLE,'UTF-8'):ucwords(strtolower($value));}
    private function dateValue(int $year,int $month,int $day):DateTimeImmutable{return(new DateTimeImmutable(sprintf('%04d-01-01',$year)))->modify(($month-1).' months')->modify(($day-1).' days');}
    /** @param list<mixed> $args */
    private function shiftMonth(array $args,bool $end):string{$date=new DateTimeImmutable($this->toText($args[0]??'now'));$date=$date->modify(((int)$this->number($args[1]??0)).' months');return($end?$date->modify('last day of this month'):$date)->format('Y-m-d');}
    private function dateDiffDays(mixed $start,mixed $end):int{return(int)(new DateTimeImmutable($this->toText($start)))->diff(new DateTimeImmutable($this->toText($end)))->format('%r%a');}
    /** @param list<mixed> $args */
    private function networkDays(array $args):int{$start=new DateTimeImmutable($this->toText($args[0]??'today'));$end=new DateTimeImmutable($this->toText($args[1]??'today'));if($start>$end)[$start,$end]=[$end,$start];$holidays=array_flip(array_map(fn($v)=>(new DateTimeImmutable($this->toText($v)))->format('Y-m-d'),$this->flatten(array_slice($args,2))));$days=0;for($date=$start;$date<=$end;$date=$date->modify('+1 day'))if((int)$date->format('N')<6&&!isset($holidays[$date->format('Y-m-d')]))$days++;return$days;}

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
