<?php
declare(strict_types=1);
$root=dirname(__DIR__); $core=dirname($root).'/mnb-phpexcel-core/src/';
spl_autoload_register(static function(string $class) use($root,$core): void {
 $prefix='Mnb\\PHPExcel\\'; if(!str_starts_with($class,$prefix)) return;
 $rel=str_replace('\\',DIRECTORY_SEPARATOR,substr($class,strlen($prefix))).'.php';
 foreach([$root.'/src/'.$rel,$core.$rel] as $p) if(is_file($p)) { require $p; return; }
});
use Mnb\PHPExcel\Format\Xlsx;
$m=Xlsx::meta(__FILE__)->quick()->only(['file']);
if(!$m instanceof Mnb\PHPExcel\Metadata\MetadataFacade) throw new RuntimeException('facade');
echo "Xlsx MetadataFacade V2 API passed.\n";
