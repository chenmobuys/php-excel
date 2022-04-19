<?php

namespace ExcelTests\Feature;


use Excel\Exception\SpreadsheetException;
use Excel\SpreadsheetFactory;
use ExcelTests\TestCase;

class SpreadsheetFactoryTest extends TestCase
{
    public function testCreateReader()
    {
        $readerCsv = SpreadsheetFactory::createReader(SpreadsheetFactory::EXCEL_CSV);
        $this->assertInstanceOf(\Excel\Reader\Csv::class, $readerCsv);

        $readerOds = SpreadsheetFactory::createReader(SpreadsheetFactory::EXCEL_ODS);
        $this->assertInstanceOf(\Excel\Reader\Ods::class, $readerOds);

        $readerXls = SpreadsheetFactory::createReader(SpreadsheetFactory::EXCEL_XLS);
        $this->assertInstanceOf(\Excel\Reader\Xls::class, $readerXls);

        $readerXlsx = SpreadsheetFactory::createReader(SpreadsheetFactory::EXCEL_XLSX);
        $this->assertInstanceOf(\Excel\Reader\Xlsx::class, $readerXlsx);

        $this->expectException(SpreadsheetException::class);
        SpreadsheetFactory::createReader('Foo');
    }

    public function testCreateReaderForFile()
    {
        $readerCsv = SpreadsheetFactory::createReaderForFile('tests/data/sample.csv');
        $this->assertInstanceOf(\Excel\Reader\Csv::class, $readerCsv);

        $readerOds = SpreadsheetFactory::createReaderForFile('tests/data/sample.ods');
        $this->assertInstanceOf(\Excel\Reader\Ods::class, $readerOds);

        $readerXls = SpreadsheetFactory::createReaderForFile('tests/data/sample.xls');
        $this->assertInstanceOf(\Excel\Reader\Xls::class, $readerXls);

        $readerXlsx = SpreadsheetFactory::createReaderForFile('tests/data/sample.xlsx');
        $this->assertInstanceOf(\Excel\Reader\Xlsx::class, $readerXlsx);

        $this->expectException(SpreadsheetException::class);
        SpreadsheetFactory::createReaderForFile('tests/data/notexists.csv');
    }

    public function testRegisterReader()
    {
        SpreadsheetFactory::registerReader('Foo', FooReader::class);
        $readerFoo = SpreadsheetFactory::createReader('Foo');
        $this->assertInstanceOf(FooReader::class, $readerFoo);
    }
}
