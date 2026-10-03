<?php
declare(strict_types=1);
$root=dirname(__DIR__);$core=dirname($root).'/mnb-phpexcel-core/src/';
spl_autoload_register(static function(string $class)use($root,$core):void{
 $prefix='Mnb\\PHPExcel\\';if(!str_starts_with($class,$prefix))return;
 $rel=str_replace('\\',DIRECTORY_SEPARATOR,substr($class,strlen($prefix))).'.php';
 foreach([$root.'/src/'.$rel,$core.$rel] as $p)if(is_file($p)){require $p;return;}
});
use Mnb\PHPExcel\Format\Xlsx;
$m=Xlsx::cloud()->account('local-test',['provider'=>'local']);
$src=tempnam(sys_get_temp_dir(),'mnbc');file_put_contents($src,'test');$dst=$src.'.copy';
$f=$m->upload($src,['path'=>$dst]);
if($f->provider!=='local'||!is_file($dst))throw new RuntimeException('cloud');
@unlink($src);@unlink($dst);echo "Xlsx cloud ecosystem passed\n";
