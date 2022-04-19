<?php

namespace ExcelTests\Feature\Shared;

use Excel\Shared\Style;
use Excel\Shared\Style\Format;
use Excel\Shared\Style\SST;
use Excel\Shared\Style\XF;
use ExcelTests\TestCase;

class StyleTest extends TestCase
{
    public function testGetXF()
    {
        $style = new Style();
        $value = $style->getXF();
        $this->assertInstanceOf(XF::class, $value);
    }

    public function testGetSST()
    {
        $style = new Style();
        $value = $style->getSST();
        $this->assertInstanceOf(SST::class, $value);
    }

    public function testGetFormat()
    {
        $style = new Style();
        $value = $style->getFormat();
        $this->assertInstanceOf(Format::class, $value);
    }
}
