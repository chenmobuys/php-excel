<?php

namespace Excel\Shared;

use Excel\Reader\CellIterator;

class Row
{
    /**
     * @var int
     */
    private $rowIndex;

    /**
     * @var \Excel\Shared\Style
     */
    private $style;

    /**
     * @var \Excel\Shared\Cell[]
     */
    private $cells;


    public function __construct(int $rowIndex, Style $style, array $cells = [])
    {
        $this->rowIndex = $rowIndex;
        $this->style = $style;
        $this->cells = $cells;
    }

    public function getRowIndex(): int
    {
        return $this->rowIndex;
    }

    public function getStyle(): Style
    {
        return $this->style;
    }

    public function getCells(): array
    {
        return array_map(function (Cell $cell) {
            return $cell->setStyle($this->style);
        }, $this->cells);
    }

    public function getCellIterator(): CellIterator
    {
        return new CellIterator($this);
    }
}
