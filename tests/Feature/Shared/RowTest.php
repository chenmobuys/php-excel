<?php

namespace ExcelTests\Feature\Shared;

use Excel\Reader\CellIterator;
use Excel\Shared\Row;
use Excel\Shared\Style;
use ExcelTests\TestCase;

class RowTest extends TestCase
{
    public function testGetRowIndex()
    {
        $style = new Style();
        $row = new Row(0, $style);

        $rowIndex = $row->getRowIndex();
        $this->assertEquals(0, $rowIndex);
    }

    public function testGetStyle()
    {
        $style = new Style();
        $row = new Row(0, $style);

        $style = $row->getStyle();
        $this->assertInstanceOf(Style::class, $style);
    }

    public function testGetCells()
    {
        $style = new Style();
        $row = new Row(0, $style);

        $cells = $row->getCells();
        $this->assertEquals([], $cells);
    }

    public function testGetCellIterator()
    {
        $style = new Style();
        $row = new Row(0, $style);

        $cellIterator = $row->getCellIterator();
        $this->assertInstanceOf(CellIterator::class, $cellIterator);
    }
}
