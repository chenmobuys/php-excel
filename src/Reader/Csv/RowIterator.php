<?php

namespace Excel\Reader\Csv;

use Excel\Reader\BaseRowIterator;
use Excel\Shared\Cell;
use Excel\Shared\Row;
use Excel\Shared\Style;
use SplFileObject;

class RowIterator extends BaseRowIterator
{
    /**
     * @var \SplFileObject
     */
    private $worksheet;

    /**
     * RowIterator constructor.
     * 
     * @param \SplFileObject $worksheet
     * @param \Excel\Shared\Style $style
     * @param int $worksheetIndex
     * @param int $startRow
     * @param int $endRow
     */
    public  function __construct(SplFileObject $worksheet, Style $style, $worksheetIndex, int $startRow = 1, int $endRow = null)
    {
        parent::__construct($style, $worksheetIndex, $startRow, $endRow);
        $this->worksheet = $worksheet;
    }

    /**
     * Get current row.
     * 
     * @return \Excel\Shared\Row
     */
    public function current(): Row
    {
        $this->row = [];
        foreach ((array) $this->worksheet->current() as $columnIndex => $cellValue) {
            $cell = new Cell($this->position, $columnIndex);
            $this->row[] = $cell->setValue($cellValue);
        }

        return parent::current();
    }

    /**
     * Rewind iterator.
     * 
     * @return void
     */
    public function rewind(): void
    {
        while (true) {
            if ($this->position >= $this->startRow - 1) {
                break;
            }

            $this->worksheet->next();
            $this->position++;
        }
    }


    /**
     * Set next position.
     * 
     * @return void
     */
    public function next(): void
    {
        parent::next();
        $this->worksheet->next();
    }
}
