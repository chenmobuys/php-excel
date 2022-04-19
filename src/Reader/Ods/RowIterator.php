<?php

namespace Excel\Reader\Ods;

use Excel\Reader\BaseRowIterator;
use Excel\Shared\Cell;
use Excel\Shared\Row;
use Excel\Shared\Style;
use XMLReader;

class RowIterator extends BaseRowIterator
{
    /**
     * @var \XMLReader
     */
    private $worksheet;

    /**
     * @var bool
     */
    private $tableOpen;

    /**
     * @var bool
     */
    private $rowOpen;

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
        $this->row = [];

        $worksheetIndex = 0;
        while (!$this->tableOpen && $this->worksheet->read()) {
            switch ($this->worksheet->name) {
                case 'table:table':
                    if ($worksheetIndex === $this->worksheetIndex) {
                        $this->tableOpen = true;
                        break 2;
                    }
                    if (
                        $worksheetIndex < $this->worksheetIndex
                        && $this->worksheet->nodeType === XMLReader::ELEMENT
                    ) {
                        $worksheetIndex++;
                    }
                    break;
            }
        }

        while ($this->tableOpen && !$this->rowOpen && $this->worksheet->read()) {
            switch ($this->worksheet->name) {
                case 'table:table-row':
                    if ($this->worksheet->nodeType === XMLReader::ELEMENT) {
                        $this->rowOpen = true;
                        break 2;
                    }
                    break;
            }
        }

        $lastCell = null;
        $paragraphs = [];
        $columnIndex = 0;
        while ($this->rowOpen && $this->worksheet->read()) {
            switch ($this->worksheet->name) {
                case 'table:table-cell':
                    if ($this->worksheet->isEmptyElement) {
                        $this->row[$columnIndex] = new Cell($this->position, $columnIndex);
                        $columnIndex++;
                        break;
                    }
                    if ($this->worksheet->nodeType == XMLReader::ELEMENT) {
                        $paragraphs = [];
                    }
                    if ($this->worksheet->nodeType == XMLReader::END_ELEMENT) {
                        $valueType = $this->worksheet->getAttribute('office:value-type');
                        $formulaValue = $this->worksheet->getAttribute('table:formula');
                        if (in_array($valueType, ['date', 'time'])) {
                            $value = $this->worksheet->getAttribute('office:' . $valueType . '-value');
                        } else {
                            $value = $this->worksheet->getAttribute('office:value') ?: implode("\n", $paragraphs);
                        }
                        $formattedValue = implode("\n", $paragraphs);
                        $lastCell = new Cell($this->position, $columnIndex);
                        $this->row[$columnIndex] = $lastCell->setValue($value)->setFormattedValue($formattedValue)->setFormulaValue($formulaValue);

                        if ($this->worksheet->getAttribute('table:number-columns-repeated') !== null) {
                            $repeatedColumnCount = $this->worksheet->getAttribute('table:number-columns-repeated');
                            // Checking if larger than one because the value is already added to the row once before
                            if ($repeatedColumnCount > 1) {
                                $this->row = array_pad($this->row, count($this->row) + $repeatedColumnCount - 1, $lastCell ?: new Cell($this->position, $columnIndex));
                            }
                        }
                        $columnIndex++;
                    }
                    break;
                case 'text:p':
                    if ($this->worksheet->nodeType != XMLReader::END_ELEMENT) {
                        $paragraphs[] = $this->worksheet->readString();
                    }
                    break;
                case 'table:table-row':
                    $this->rowOpen = false;
                    break 2;
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
        $worksheetIndex = 0;
        while (!$this->tableOpen && $this->worksheet->read()) {
            switch ($this->worksheet->name) {
                case 'table:table':
                    if ($worksheetIndex == $this->worksheetIndex) {
                        $this->tableOpen = true;
                        break 2;
                    }
                    if (
                        $worksheetIndex < $this->worksheetIndex
                        && $this->worksheet->nodeType == XMLReader::ELEMENT
                    ) {
                        $worksheetIndex++;
                    }
                    break;
            }
        }

        while ($this->startRow > 1 && $this->tableOpen && !$this->rowOpen && $this->worksheet->read()) {
            switch ($this->worksheet->name) {
                case 'table:table-row':
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
