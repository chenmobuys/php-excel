<?php

namespace Excel\Reader\Xlsx;

use XMLReader;
use Excel\Shared\Cell;
use Excel\Shared\Row;
use Excel\Shared\Style;
use Excel\Shared\Coordinate;
use Excel\Reader\BaseRowIterator;

class RowIterator extends BaseRowIterator
{
    /**
     * @var XMLReader
     */
    private $worksheet;

    /**
     * RowIterator constructor.
     * 
     * @param \XMLReader $worksheet
     * @param \Excel\Shared\Style $style
     * @param int $worksheetIndex
     * @param int $startRow
     * @param int $endRow
     */
    public function __construct(XMLReader $worksheet, Style $style, int $worksheetIndex, int $startRow = 1, int $endRow = null)
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
        $rowOpen = false;
        $columnIndex = 0;
        while ($this->worksheet->read()) {
            if ($this->worksheet->name == 'row') {
                // Getting the row spanning area (stored as e.g., 1:12)
                // so that the last cells will be present, even if empty
                $rowSpans = $this->worksheet->getAttribute('spans');
                if ($rowSpans) {
                    $rowSpans = explode(':', $rowSpans);
                    $rowColumnCount = $rowSpans[1];
                } else {
                    $rowColumnCount = 0;
                }

                if ($rowColumnCount > 0) {
                    $this->row = array_fill(0, $rowColumnCount, new Cell($this->position, $columnIndex));
                }
                $columnIndex++;
                $rowOpen = true;
                break;
            }
        }


        if (!$rowOpen) {
            return $this->row;
        }

        // These two are needed to control for empty cells
        $index = -1;
        $styleId = -1;
        $maxIndex = 0;
        $cellCount = 0;
        $columnIndex = 0;
        $cellHasSharedString = false;

        while ($rowOpen && $this->worksheet->read()) {
            switch ($this->worksheet->name) {
                    // End of row
                case 'row':
                    if ($this->worksheet->nodeType == XMLReader::END_ELEMENT) {
                        break 2;
                    }
                    break;
                    // Cell
                case 'c':
                    // If it is a closing tag, skip it
                    if ($this->worksheet->nodeType == XMLReader::END_ELEMENT) {
                        $columnIndex++;
                        break;
                    }

                    // Get the index of the cell
                    $index = $this->worksheet->getAttribute('r');
                    $letter = preg_replace('{[^[:alpha:]]}S', '', $index);
                    $index = Coordinate::columnIndexFromColumnLetter($letter);

                    // Get the style of the cell
                    $styleId = (int) $this->worksheet->getAttribute('s');

                    // Determine cell type
                    if ($this->worksheet->getAttribute('t') == 's') {
                        $cellHasSharedString = true;
                    } else {
                        $cellHasSharedString = false;
                    }

                    $cellCount++;
                    if ($index > $maxIndex) {
                        $maxIndex = $index;
                    }

                    break;
                    // Cell value
                case 'v':
                case 'is':
                    if ($this->worksheet->nodeType != XMLReader::END_ELEMENT) {
                        $value = $this->worksheet->readString();
                        if ($index >= 0) {
                            $cell = new Cell($this->position, $columnIndex);
                            $cellHasSharedString ? $cell->setSstIndex($value) : $cell->setValue($value);
                            $this->row[$index] = $cell->setXfIndex($styleId);
                        }
                    }
                    break;
                    // Formula value
                case 'f':
                    if ($this->worksheet->nodeType != XMLReader::END_ELEMENT) {
                        $value = $this->worksheet->readString();
                        if ($this->row[$index] instanceof Cell) {
                            $this->row[$index]->setFormulaValue($value);
                        }
                    }

                    break;
            }
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
        if ($this->startRow > 1) {
            while ($this->worksheet->read()) {
                switch ($this->worksheet->name) {
                        // End of row
                    case 'row':
                        if ($this->worksheet->nodeType == XMLReader::ELEMENT) {
                            $this->position++;
                        }

                        if ($this->position >= $this->startRow - 1) {
                            break 2;
                        }

                        break;
                }
            }
        }
    }
}
