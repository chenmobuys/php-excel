<?php

namespace ExcelTests\Feature\Reader;

use Excel\Reader\CellIterator;
use Excel\Shared\Cell;
use Excel\Shared\Row;
use Excel\Shared\Style;
use ExcelTests\TestCase;

class CellIteratorTest extends TestCase
{
    const CELLS = [
        110, 120, 130,
        111, 121, 131,
        112, 123, 133,
    ];

    public function testIterator()
    {
        $cells = [];
        foreach (self::CELLS as $index => $cell) {
            $cells[] = (new Cell(0, $index))->setValue($cell);
        }

        $row = new Row(0, new Style, $cells);
        $cellIterator = new CellIterator($row);
        $this->assertEquals($cellIterator->count(), count(self::CELLS));

        $rowValue = [];
        foreach ($cellIterator as $cell) {
            $rowValue[] = $cell->getValue();
        }

        $this->assertEquals($rowValue, self::CELLS);
    }
}
