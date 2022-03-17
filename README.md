# php_excel_export
PHPExcel from Array by cli or browser, support php5, php7, php8

通过命令行或者浏览器把PHP数组中的数据导出到excel中。同时支持php低版本和高版本

## Use example: 使用示例

### 修改 项目中 `composer.json` 文件，如下: 重点是加上 `require` 和 `repositories`
> vi composer.json
> 
```{
    "name": "test",
    "type": "library",
    
    "minimum-stability": "dev",
    "require": {
        "jacena/php-excel-export": "dev-master"
    },
    "repositories": {
        "jacena/php-excel-export": {
            "type": "git",
            "url": "https://github.com/jacena/php-excel-export.git"
        }
    }
}
```

### 必须加上  `--ignore-platform-reqs` 否则有些有些包装不上
- 安装 install 
    > composer install --ignore-platform-reqs -vvv
- 更新 update
    > composer require --ignore-platform-reqs -vvv

### 调用示例 demo

> vi test.php

> 
```
<?php
namespace test;

require_once("vendor/autoload.php"); // 如果用的框架，肯定已经自动加载了。可以注释

use Jacena\PhpExcelMaker\PHPExcelMaker;

$keys = $title = [];


$data = [
    ['aaa'=>'dkdk', 'bbb'=>'ddkdkd', 'ccc'=>'dkdkfd'],
    ['aaa'=>'dkdk', 'bbb'=>'ddkdkd', 'ccc'=>'dkdkfd'],
    ['aaa'=>'dkdk', 'bbb'=>'ddkdkd', 'ccc'=>'dkdkfd'],
    ['aaa'=>'dkdk', 'bbb'=>'ddkdkd', 'ccc'=>'dkdkfd'],
    ['aaa'=>'dkdk', 'bbb'=>'ddkdkd', 'ccc'=>'dkdkfd'],
];


$title = [
    'aaa' => '姓名',
    'bbb' => '年龄',
    'ccc' => '性别',
];

$excel = new PHPExcelMaker();


if (PHP_SAPI == 'cli') {
    
    // var_dump($excel->getPHPVersion());exit;
    $excel->exportExcel($keys, $title, $data, 'xxx'); // 数组名和文件名一致
    exit;

}else {
    
    // var_dump($excel->getPHPVersion());exit;
    $excel->exportExcel($keys, $title, $data, 'xxxx', false); // 数组名和文件名一致
}

```



