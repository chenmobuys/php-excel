<?php

namespace ExcelTests\Feature\Reader\Csv;

use Excel\Reader\Csv\Delimiter;
use ExcelTests\TestCase;
use SplFileObject;

class DelimiterTest extends TestCase
{
    const TEST_FILENAME = 'tests/data/sample.csv';

    public function testGetDefaultDelimiter()
    {
        $splFileObject = new SplFileObject(self::TEST_FILENAME);
        $delimiter = new Delimiter($splFileObject, '"', '\\');

        $this->assertEquals(',', $delimiter->getDefaultDelimiter());
    }

    public function testLinesCounted()
    {
        $splFileObject = new SplFileObject(self::TEST_FILENAME);
        $delimiter = new Delimiter($splFileObject, '"', '\\');

        $this->assertEquals(1000, $delimiter->linesCounted());
    }
}
