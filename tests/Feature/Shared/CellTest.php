<?php

namespace ExcelTests\Feature\Shared;

use Excel\Shared\Cell;
use Excel\Shared\Style;
use ExcelTests\TestCase;

class CellTest extends TestCase
{
    public function testRowIndex()
    {
        $cell = new Cell(0, 0);
        $rowIndex = $cell->getRowIndex();
        $this->assertEquals(0, $rowIndex);
    }

    public function testColumnIndex()
    {
        $cell = new Cell(0, 0);
        $columnIndex = $cell->getColumnIndex();
        $this->assertEquals(0, $columnIndex);
    }

    public function testColumnLetter()
    {
        $cell = new Cell(0, 0);
        $columnLetter = $cell->getColumnLetter();
        $this->assertEquals('A', $columnLetter);
    }

    public function testCoordinate()
    {
        $cell = new Cell(0, 0);
        $coordinate = $cell->getCoordinate();
        $this->assertEquals('A1', $coordinate);
    }

    public function testValue()
    {
        $cell = new Cell(0, 0);
        $style = new Style();
        $cell->setStyle($style);

        $cellNew = $cell->setValue('Foo');
        $this->assertInstanceOf(Cell::class, $cellNew);

        $value = $cell->getValue();
        $this->assertEquals('Foo', $value);
    }

    public function testFormattedValue()
    {
        $cell = new Cell(0, 0);
        $style = new Style();
        $cell->setStyle($style);

        $cellNew = $cell->setFormattedValue('Foo');
        $this->assertInstanceOf(Cell::class, $cellNew);

        $value = $cell->getFormattedValue();
        $this->assertEquals('Foo', $value);
    }

    public function testStyle()
    {
        $cell = new Cell(0, 0);
        $style = new Style();

        $cellNew = $cell->setStyle($style);
        $this->assertInstanceOf(Cell::class, $cellNew);

        $style = $cell->getStyle();
        $this->assertInstanceOf(Style::class, $style);
    }

    public function testXfIndex()
    {
        $cell = new Cell(0, 0);

        $cellNew = $cell->setXfIndex('Foo');
        $this->assertInstanceOf(Cell::class, $cellNew);

        $value = $cell->getXfIndex();
        $this->assertEquals('Foo', $value);
    }

    public function testSstIndex()
    {
        $cell = new Cell(0, 0);

        $cellNew = $cell->setSstIndex('Foo');
        $this->assertInstanceOf(Cell::class, $cellNew);

        $value = $cell->getSstIndex();
        $this->assertEquals('Foo', $value);
    }
}
