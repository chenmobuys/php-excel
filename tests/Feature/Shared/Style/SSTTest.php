<?php

namespace ExcelTests\Feature\Shared\Style;

use Excel\Shared\Cell;
use Excel\Shared\Style;
use Excel\Shared\Style\SST;
use ExcelTests\TestCase;

class SSTTest extends TestCase
{
    public function testAppendAndAll()
    {
        $style = new Style();
        $sst = new SST($style);
        $sst->append('Foo', 'Bar');
        $value = $sst->all();
        $this->assertEquals(['Foo' => 'Bar'], $value);
    }

    public function testGetValue()
    {
        $cell = new Cell(0, 0);
        $style = new Style();
        $sst = new SST($style);
        $sst->append('Foo', 'Bar');
        $style->getSST()->append('Foo', 'Bar');
        $cell->setSstIndex('Foo');
        $cell->setStyle($style);

        $value = $sst->getValue($cell);
        $this->assertEquals('Bar', $value);
    }
}
