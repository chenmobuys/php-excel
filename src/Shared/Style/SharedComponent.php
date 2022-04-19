<?php

namespace Excel\Shared\Style;

use Excel\Shared\Cell;
use Excel\Shared\Style;

abstract class SharedComponent
{
    /**
     * Items
     * 
     * @var array
     */
    protected $items = [];

    /**
     * Style
     * 
     * @var \Excel\Shared\Style
     */
    protected $style;

    /**
     * SharedComponent class constructor.
     * 
     * @param \Excel\Shared\Style $style
     */
    public function __construct(Style $style)
    {
        $this->style = $style;
    }

    /**
     * Append to items.
     * 
     * @param string $key
     * @param string $val
     * @return void
     */
    public function append(string $key, string $val = null): void
    {
        if (func_num_args() > 1) {
            $this->items[$key] = $val;
        } else {
            $this->items[] = $key;
        }
    }

    /**
     * Get all items.
     * 
     * @return array
     */
    public function all(): array
    {
        return $this->items;
    }

    /**
     * Get value.
     * 
     * @param \Excel\Shared\Cell
     * @return string
     */
    public function getValue(Cell $cell): string
    {
        return $this->items[$cell->getXfIndex()] ?? null;
    }
}
