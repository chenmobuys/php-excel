<?php

namespace ExcelTests\Feature\Shared\Style;

use Excel\Shared\Cell;
use Excel\Shared\Style;
use Excel\Shared\Style\XF;
use ExcelTests\TestCase;

class XFTest extends TestCase
{
    public function testAppendAndAll()
    {
        $style = new Style();
        $xf = new XF($style);
        $xf->append('Foo', 'Bar');
        $value = $xf->all();
        $this->assertEquals(['Foo' => 'Bar'], $value);
    }

    public function testGetValue()
    {
        $cell = new Cell(0, 0);
        $style = new Style();      
        $xf = new XF($style);
        $xf->append('Foo', 'Bar');
        $style->getXF()->append('Foo', 'Bar');
        $cell->setXfIndex('Foo');
        $cell->setStyle($style);

        $value = $xf->getValue($cell);
        $this->assertEquals('Bar', $value);
    }
}
