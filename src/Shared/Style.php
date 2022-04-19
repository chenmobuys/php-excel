<?php

namespace Excel\Shared;

use Excel\Shared\Style\XF;
use Excel\Shared\Style\SST;
use Excel\Shared\Style\Format;

class Style
{
    /**
     * @var \Excel\Shared\Style\XF
     */
    private $xf;

    /**
     * @var \Excel\Shared\Style\SST
     */
    private $sst;

    /**
     * @var \Excel\Shared\Style\Format
     */
    private $format;

    public function __construct()
    {
        $this->xf = new XF($this);
        $this->sst = new SST($this);
        $this->format = new Format($this);
    }

    public function getXF(): XF
    {
        return $this->xf;
    }

    public function getSST(): SST
    {
        return $this->sst;
    }

    public function getFormat(): Format
    {
        return $this->format;
    }
}
