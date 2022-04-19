<?php

namespace Excel\Reader;

use Countable;
use Iterator;
use Excel\Shared\Cell;
use Excel\Shared\Row;

class CellIterator implements Iterator, Countable
{
    /**
     * Cells collection
     * 
     * @var \Excel\Shared\Cell[]
     */
    private $cells;

    /**
     * Cells' count
     */
    private $count;

    /**
     * Iterator position
     */
    private $position = 0;

    /**
     * CellIterator contructor.
     * 
     * @param \Excel\Shared\Row
     */
    public function __construct(Row $row)
    {
        $this->cells = $row->getCells();
        $this->count = count($this->cells);
    }

    /**
     * Get current cell.
     * 
     * @return \Excel\Shared\Cell
     */
    public function current(): Cell
    {
        return $this->cells[$this->position];
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
        return $this->position < $this->count;
    }

    /**
     * Get cell count.
     * 
     * @return int
     */
    public function count(): int
    {
        return $this->count;
    }
}
