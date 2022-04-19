<?php

namespace Excel\Shared;

class Cell
{
    /**
     * Row index
     * 
     * @var int $rowIndex
     */
    private $rowIndex;

    /**
     * Column index
     * 
     * @var int $columnIndex
     */
    private $columnIndex;

    /**
     * XF index
     * 
     * @var int $xfIndex
     */
    private $xfIndex;

    /**
     * SST index
     * 
     * @var int $sstIndex
     */
    private $sstIndex;

    /**
     * Cell value
     * 
     * @var string
     */
    private $value;

    /**
     * Formula value
     * 
     * @var string 
     */
    private $formulaValue;

    /**
     * Cell formatted value
     * 
     * @var string
     */
    private $formattedValue;

    /**
     * Cell's style
     * 
     * @var \Excel\Shared\Style
     */
    private $style;

    /**
     * Cell class constructor.
     * 
     * @param int $rowIndex
     * @param int $columnIndex
     */
    public function __construct(int $rowIndex, int $columnIndex)
    {
        $this->rowIndex = $rowIndex;
        $this->columnIndex = $columnIndex;
    }

    public function getRowIndex()
    {
        return $this->rowIndex;
    }

    public function getColumnIndex()
    {
        return $this->columnIndex;
    }

    public function getColumnLetter()
    {
        return Coordinate::columnLetterFromColumnIndex($this->columnIndex);
    }

    public function getCoordinate()
    {
        return $this->getColumnLetter() . ($this->getRowIndex() + 1);
    }

    public function getValue()
    {
        return $this->getStyle()->getSST()->getValue($this) ?: $this->value;
    }

    public function getFormulaValue()
    {
        return $this->formulaValue;
    }

    public function getFormattedValue()
    {
        return $this->formattedValue ?: $this->getStyle()->getFormat()->getValue($this);
    }

    public function getStyle()
    {
        return $this->style;
    }

    public function getXfIndex()
    {
        return $this->xfIndex;
    }

    public function getSstIndex()
    {
        return $this->sstIndex;
    }

    public function setValue($value)
    {
        $this->value = $value;
        return $this;
    }

    public function setFormulaValue($formulaValue)
    {
        $this->formulaValue = $formulaValue;
        return $this;
    }

    public function setFormattedValue($formattedValue)
    {
        $this->formattedValue = $formattedValue;
        return $this;
    }

    public function setStyle(Style $style)
    {
        $this->style = $style;
        return $this;
    }

    public function setXfIndex($xfIndex)
    {
        $this->xfIndex = $xfIndex;
        return $this;
    }

    public function setSstIndex($sstIndex)
    {
        $this->sstIndex = $sstIndex;
        return $this;
    }

    public function __toString()
    {
        return $this->getValue();
    }
}
