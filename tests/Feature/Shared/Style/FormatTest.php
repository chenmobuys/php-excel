<?php

namespace ExcelTests\Feature\Shared\Style;

use Excel\Shared\Cell;
use Excel\Shared\Style;
use Excel\Shared\Style\Format;
use ExcelTests\TestCase;

class FormatTest extends TestCase
{
    public function testAppendAndAll()
    {
        $style = new Style();
        $format = new Format($style);
        $format->append('Foo', 'Bar');
        $value = $format->all();
        $this->assertEquals(['Foo' => 'Bar'], $value);
    }

    public function testGetValue()
    {
        $cell = new Cell(0, 0);
        $style = new Style();
        $cell->setValue('Foo');
        $cell->setStyle($style);
        $format = new Format($style);

        $value = $format->getValue($cell);
        $this->assertEquals('Foo', $value);
    }
}
