<?php

namespace ExcelTests\Feature\Reader;

use Excel\Exception\SpreadsheetException;
use Excel\Reader\BaseRowIterator;
use Excel\SpreadsheetFactory;
use ExcelTests\TestCase;

class CsvTest extends TestCase
{
    const TEST_FILENAME = 'tests/data/sample.csv';

    public function testLoad()
    {
        $reader = SpreadsheetFactory::createReaderForFile(self::TEST_FILENAME);
        $reader->load(self::TEST_FILENAME);

        $worksheetNames = $reader->getWorksheetNames();
        $this->assertEquals(['sample'], $worksheetNames);

        $worksheetInfo = $reader->getWorksheetInfo();
        $worksheetInfoExpected = [
            [
                'worksheetName' => 'sample',
                'lastColumnLetter' => 'C',
                'lastColumnIndex' => 2,
                'totalRows' => 2,
                'totalColumns' => 3,
            ]
        ];
        $this->assertEquals($worksheetInfoExpected, $worksheetInfo);
    }

    public function testIsReadable()
    {
        $reader = SpreadsheetFactory::createReaderForFile(self::TEST_FILENAME);
        $actual = $reader->isReadable(self::TEST_FILENAME);
        $this->assertTrue($actual);

        $this->expectException(SpreadsheetException::class);
        $reader = SpreadsheetFactory::createReaderForFile('notexists_' . self::TEST_FILENAME);
        $actual = $reader->isReadable(self::TEST_FILENAME);
    }

    public function testListWorksheetNames()
    {
        $reader = SpreadsheetFactory::createReaderForFile(self::TEST_FILENAME);
        $worksheetNames = $reader->listWorksheetNames(self::TEST_FILENAME);
        $this->assertEquals(['sample'], $worksheetNames);
    }

    public function testListWorksheetInfo()
    {
        $reader = SpreadsheetFactory::createReaderForFile(self::TEST_FILENAME);
        $worksheetInfo = $reader->listWorksheetInfo(self::TEST_FILENAME);
        $worksheetInfoExpected = [
            [
                'worksheetName' => 'sample',
                'lastColumnLetter' => 'C',
                'lastColumnIndex' => 2,
                'totalRows' => 2,
                'totalColumns' => 3,
            ]
        ];
        $this->assertEquals($worksheetInfoExpected, $worksheetInfo);
    }

    public function testGetRowIterator()
    {
        $reader = SpreadsheetFactory::createReaderForFile(self::TEST_FILENAME);
        $reader->load(self::TEST_FILENAME);

        $this->expectException(SpreadsheetException::class);
        $reader->getRowIteratorByWorksheetName('Foo');

        $rowIterator = $reader->getRowIteratorByWorksheetName('sample');
        $this->assertInstanceOf(BaseRowIterator::class, $rowIterator);
        $actualRows = [];
        foreach ($rowIterator as $row) {
            foreach ($row->getCellIterator() as $cell) {
                $actualRows[$row->getRowIndex()][$cell->getColumnIndex()] = $cell->getValue();
            }
        }
        $expectedRows = [
            1 => ['Title1', 'Title2', 'Title3'],
            2 => ['Desc1', 'Desc2', 'Desc3'],
        ];

        $this->assertEquals($expectedRows, $actualRows);
    }

    public function testGetActiveRowIterator()
    {
        $reader = SpreadsheetFactory::createReaderForFile(self::TEST_FILENAME);
        $reader->load(self::TEST_FILENAME);
        $rowIterator = $reader->getActiveRowIterator();
        $this->assertInstanceOf(BaseRowIterator::class, $rowIterator);
    }

    public function testGetRowIterators()
    {
        $reader = SpreadsheetFactory::createReaderForFile(self::TEST_FILENAME);
        $reader->load(self::TEST_FILENAME);
        $rowIterators = $reader->getRowIterators();
        $this->assertEquals(1, count($rowIterators));
        $this->assertInstanceOf(BaseRowIterator::class, current($rowIterators));
    }
}
