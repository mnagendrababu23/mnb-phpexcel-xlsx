<?php

declare(strict_types=1);

namespace Mnb\PHPExcel\Reader\Formula;

use Mnb\PHPExcel\Support\MnbExcelException;

/** Instance-scoped custom formula functions for application-specific formulas. */
final class FormulaFunctionRegistry
{
    /** @var array<string,callable(list<mixed>):mixed> */
    private array $functions=[];
    /** @param callable(list<mixed>):mixed $function */
    public function register(string $name,callable $function):self{$name=strtoupper(trim($name));if(preg_match('/^[A-Z_][A-Z0-9_.]*$/',$name)!==1)throw new MnbExcelException('Invalid formula function name: '.$name);$this->functions[$name]=$function;return $this;}
    public function has(string $name):bool{return isset($this->functions[strtoupper($name)]);}
    /** @param list<mixed> $args */
    public function call(string $name,array $args):mixed{if(!$this->has($name))throw new MnbExcelException('Formula function is not registered: '.$name);return ($this->functions[strtoupper($name)])($args);}
    /** @return list<string> */
    public function names():array{$names=array_keys($this->functions);sort($names);return $names;}
}
