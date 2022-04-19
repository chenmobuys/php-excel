<?php

namespace ExcelTests\Feature\Shared;

use Excel\Shared\CodePage;
use ExcelTests\TestCase;

class CodePageTest extends TestCase
{
    public function testNumberToName()
    {
        $codePage = CodePage::numberToName(65001);
        $this->assertEquals('UTF-8', $codePage);

        $codePage = CodePage::numberToName(0);
        $this->assertEquals(CodePage::DEFUALT_CODE_PAGE, $codePage);
    }
}
