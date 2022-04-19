<?php

namespace ExcelTests\Feature\Shared;

use Excel\Shared\Coordinate;
use ExcelTests\TestCase;

class CoordinateTest extends TestCase
{
    public function testColumnIndexFromColumnLetter()
    {
        $columnIndex = Coordinate::columnIndexFromColumnLetter('A');
        $this->assertEquals(0,  $columnIndex);
    }

    public function testColumnLetterFromColumnIndex()
    {
        $columnLetter = Coordinate::columnLetterFromColumnIndex(0);
        $this->assertEquals('A',  $columnLetter);
    }
}
