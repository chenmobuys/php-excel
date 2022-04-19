<?php

namespace ExcelTests\Feature\Reader\Xls;

use Excel\Reader\Xls\OLEReader;
use ExcelTests\TestCase;

class OLEReaderTest extends TestCase
{
    const TEST_FILENAME = 'tests/data/sample.xls';

    public function testIsReadable()
    {
        $this->assertTrue(OLEReader::isReadable(self::TEST_FILENAME));

        $this->assertFalse(OLEReader::isReadable('notexists_' . self::TEST_FILENAME));
    }

    public function testGetWorksheets()
    {
        $OLEReader = new OLEReader(self::TEST_FILENAME);
        $actualWorksheets = $OLEReader->getWorksheets();
        $expectedWorksheets = [
            [
                'name' => 'Sheet1',
                'type' => 0,
                'offset' => 13135,
                'state' => 0,
            ],
            [
                'name' => 'Sheet2',
                'type' => 0,
                'offset' => 13671,
                'state' => 0,
            ],
            [
                'name' => 'Sheet3',
                'type' => 0,
                'offset' => 14067,
                'state' => 0,
            ]
        ];
        $this->assertEquals($expectedWorksheets, $actualWorksheets);
    }

    public function testGetWorkbook()
    {
        $OLEReader = new OLEReader(self::TEST_FILENAME);
        $workbook = $OLEReader->getWorkbook();
        $workbookLength = strlen($workbook);
        $this->assertGreaterThan(0, $workbookLength);
    }

    public function testGetWorksheetByIndex()
    {
        $OLEReader = new OLEReader(self::TEST_FILENAME);

        $expectedWorksheet = [
            'name' => 'Sheet1',
            'type' => 0,
            'offset' => 13135,
            'state' => 0
        ];
        $actualWorksheet = $OLEReader->getWorksheetByIndex(0);
        $this->assertEquals($expectedWorksheet, $actualWorksheet);
    }

    public function testGetRowOffsetsByIndex()
    {
        $OLEReader = new OLEReader(self::TEST_FILENAME);

        $expectedRowOffsetsCount = 2;
        $actualRowOffsetsCount = count($OLEReader->getRowOffsetsByIndex(0));
        $this->assertEquals($expectedRowOffsetsCount, $actualRowOffsetsCount);
    }

    public function testGetInt()
    {
        $OLEReader = new OLEReader(self::TEST_FILENAME);
        $workbook = $OLEReader->getWorkbook();
        $actualInt = $OLEReader->getInt(0, $workbook, 2);
        $this->assertEquals(0x809, $actualInt);
    }
}
