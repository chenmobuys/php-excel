<?php

namespace ExcelTests\Feature;

use Excel\Settings;
use ExcelTests\TestCase;

class SettingsTest extends TestCase
{
    public function testGetLibXmlLoaderOptions()
    {
        $libXmlLoaderOptions = Settings::getLibXmlLoaderOptions();
        $this->assertEquals(LIBXML_DTDLOAD | LIBXML_DTDATTR, $libXmlLoaderOptions);
    }

    public function testSetLibXmlLoaderOptions()
    {
        Settings::setLibXmlLoaderOptions(LIBXML_DTDLOAD);
        $libXmlLoaderOptions = Settings::getLibXmlLoaderOptions();
        $this->assertEquals(LIBXML_DTDLOAD, $libXmlLoaderOptions);

        Settings::setLibXmlLoaderOptions(LIBXML_DTDLOAD | LIBXML_DTDATTR);
        $libXmlLoaderOptions = Settings::getLibXmlLoaderOptions();
        $this->assertEquals(LIBXML_DTDLOAD | LIBXML_DTDATTR, $libXmlLoaderOptions);
    }
}
