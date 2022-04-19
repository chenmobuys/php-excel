<?php

namespace Excel\Shared\Style;

use Excel\Shared\Cell;

class SST extends SharedComponent
{
    /**
     * Get value.
     * 
     * @param \Excel\Shared\Cell
     * @return string
     */
    public function getValue(Cell $cell): string
    {
        return $this->items[$cell->getSstIndex()] ?? '';
    }
}
