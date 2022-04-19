## PHPExcel

[![Build Status](https://github.com/chenmobuys/php-excel/workflows/master/badge.svg)](https://github.com/chenmobuys/php-excel/actions)
[![Latest Stable Version](https://img.shields.io/packagist/v/chen/php-excel.svg)](https://packagist.org/packages/chen/php-excel) 
[![Total Downloads](https://img.shields.io/packagist/dt/chen/php-excel)](https://packagist.org/packages/chen/php-excel) 
[![License](https://img.shields.io/packagist/l/chen/php-excel)](https://packagist.org/packages/chen/php-excel) 
[![Platform Support](https://img.shields.io/packagist/php-v/chen/php-excel)](https://github.com/chenmobuys/php-excel)

## 描述

PHPExcel 主要目的是使用尽量少的内存读取大型文件。

## 安装

使用 `composer` 安装

```bash
composer require chen/php-excel -vvv
```

## 用法

### 初始化 Reader

#### 根据类型初始化 Reader

```php
<?php

use Excel\SpreadsheetFactory;

$reader = SpreadsheetFactory::createReader(SpreadsheetFactory::EXCEL_CSV);
```

#### 根据文件后缀名初始化 Reader

```php
<?php

use Excel\SpreadsheetFactory;

$filename = 'sample.csv';
$reader = SpreadsheetFactory::createReaderForFile($filename);
```

#### 自定义 Reader

```php
<?php

use Excel\Reader\BaseReader;

class SampleReader extends BaseReader {
    
    /**
     * Load filename
     *
     * @return void
     */
    public function load($filename): void 
    {
        // TODO
    }
    ...
}

----------------------------------------------

<?php

use Excel\SpreadsheetFactory;

SpreadsheetFactory::registerReader('Xlsx', SampleReader::class);

```

### 读取数据

#### 获取工作表名称

```php
<?php

use Excel\SpreadsheetFactory;

$filename = 'sample.xlsx';
$reader = SpreadsheetFactory::createReaderForFile($filename);

/**
 * @example
 * [
 *     'Sheet1',
 *     'Sheet2',
 *     'Sheet3',
 * ]
 */
$worksheetNames = $reader->listWorksheetNames($filename);
```

#### 获取工作表信息

```php
<?php

use Excel\SpreadsheetFactory;

$filename = 'sample.xlsx';
$reader = SpreadsheetFactory::createReaderForFile($filename);

/**
 * @example
 * [
 *     [
 *          'worksheetName' => 'Sheet1',
 *          'lastColumnLetter' => 'C', 
 *          'lastColumnIndex' => '2', 
 *          'totalRows' => '2', 
 *          'totalColumns' => '3', 
 *     ],
 *     [
 *          'worksheetName' => 'Sheet2',
 *          'lastColumnLetter' => 'C', 
 *          'lastColumnIndex' => '2', 
 *          'totalRows' => '2', 
 *          'totalColumns' => '3', 
 *     ],
 *     [
 *          'worksheetName' => 'Sheet3',
 *          'lastColumnLetter' => 'C', 
 *          'lastColumnIndex' => '2', 
 *          'totalRows' => '2', 
 *          'totalColumns' => '3', 
 *     ],
 * ]
 */
$worksheetInfo = $reader->listWorksheetInfo($filename);
```

#### 根据表索引获取行迭代器
```php
<?php

use Excel\SpreadsheetFactory;

$filename = 'sample.xlsx';
$reader = SpreadsheetFactory::createReaderForFile($filename);

foreach($reader->getRowIteratorByWorksheetIndex(0) as $row) {

    foreach($row->getCellIterator() as $cell) {

        // TODO...
    }

}

```

#### 根据表名称获取行迭代器
```php
<?php

use Excel\SpreadsheetFactory;

$filename = 'sample.xlsx';
$reader = SpreadsheetFactory::createReaderForFile($filename);

foreach($reader->getRowIteratorByWorksheetName('Sheet1') as $row) {

    foreach($row->getCellIterator() as $cell) {

        // TODO...
    }

}

```

