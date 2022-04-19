<?php

namespace Excel\Reader;

use Countable;
use Excel\Shared\Row;
use Excel\Shared\Style;
use Iterator;

abstract class BaseRowIterator implements Iterator, Countable
{
    /**
     * Current row.
     *
     * @var array
     */
    protected $row = [];

    /**
     * Current iterator position.
     *
     * @var int
     */
    protected $position = 0;

    /**
     * Start position.
     *
     * @var int
     */
    protected $startRow = 1;

    /**
     * End position.
     *
     * @var int
     */
    protected $endRow = 1;

    /**
     * Row count
     * 
     * @var int
     */
    protected $count;

    /**
     * Style
     * 
     * @var \Excel\Shared\Style
     */
    protected $style;

    /**
     * Worksheet index
     * 
     * @var int
     */
    protected $worksheetIndex;

    /**
     * BaseRowIterator constructor.
     * 
     * @param \Excel\Shared\Style
     * @param int $worksheetIndex
     * @param int $startRow
     * @param int $endRow
     */
    public function __construct(Style $style, int $worksheetIndex = 0, int $startRow = 1, int $endRow = null)
    {
        $this->style = $style;
        $this->worksheetIndex = $worksheetIndex;
        $this->startRow = $startRow;
        $this->endRow = $endRow;
    }

    /**
     * Get current row.
     * 
     * @return \Excel\Shared\Row
     */
    public function current(): Row
    {
        return new Row($this->position + 1, $this->style, $this->row);
    }

    /**
     * Get current key.
     * 
     * @return int
     */
    public function key(): int
    {
        return $this->position;
    }


    /**
     * Set next position.
     * 
     * @return void
     */
    public function next(): void
    {
        ++$this->position;
    }

    /**
     * Rewind iterator.
     * 
     * @return void
     */
    public function rewind(): void
    {
        $this->position = 0;
    }

    /**
     * Determine whether has next row.
     * 
     * @return bool
     */
    public function valid(): bool
    {
        return $this->position >= $this->startRow - 1 && $this->position < $this->endRow;
    }

    /**
     * Get row count.
     * 
     * @return int
     */
    public function count(): int
    {
        return $this->endRow - $this->startRow + 1;
    }
}
