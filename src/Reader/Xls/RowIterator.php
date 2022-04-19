<?php

namespace Excel\Reader\Xls;

use Excel\Reader\BaseRowIterator;
use Excel\Shared\Cell;
use Excel\Shared\Style;
use Excel\Reader\Xls\OLEReader;
use Excel\Shared\Row;

/**
 * @see http://www.openoffice.org/sc/excelfileformat.pdf#4.7.3
 */
class RowIterator extends BaseRowIterator
{
    /**
     * @var \Excel\Reader\Xls\OLEReader
     */
    private $worksheet;

    /** 
     * @var string
     */
    private $workbook;

    /**
     * @var string
     */
    private $worksheetVersion;

    /**
     * @var int
     */
    private $worksheetOffset;

    /**
     * @var int
     */
    private $currentOffset;

    /**
     * @var int
     */
    private $currentLength;

    /**
     * @var array
     */
    private $rowOffsets = [];

    /**
     * RowIterator constructor.
     * 
     * @param \Excel\Reader\Xls\OLEReader $worksheet
     * @param \Excel\Shared\Style $style
     * @param int $worksheetIndex
     * @param int $startRow
     * @param int $endRow
     */
    public function __construct(OLEReader $worksheet, Style $style, int $worksheetIndex, int $startRow = 1, int $endRow = null)
    {
        parent::__construct($style, $worksheetIndex, $startRow, $endRow);
        $this->worksheet = $worksheet;
        $this->workbook = $worksheet->getWorkbook();
        $this->worksheetOffset = $worksheet->getWorksheetByIndex($this->worksheetIndex)['offset'];
        $this->rowOffsets = $worksheet->getRowOffsetsByIndex($this->worksheetIndex);
    }

    /**
     * Get current row.
     * 
     * @return \Excel\Shared\Row
     */
    public function current(): Row
    {
        $this->row = [];

        // 4.7.3 Cell Block
        do {
            $position = $this->currentOffset + $this->currentLength;

            $lowCode = $this->worksheet->getInt($position, $this->workbook, 1);
            if ($lowCode == OLEReader::XLS_WORKSHEET_EOF) {
                break;
            }

            $code = $this->worksheet->getInt($position, $this->workbook, 2);
            $length = $this->worksheet->getInt($position + 2, $this->workbook, 2);
            $position += 4;

            $currentRowOffset = $this->rowOffsets[$this->position + 1];
            $nextRowOffset = $this->rowOffsets[$this->position + 2] ?? null;
            // Determine whether next row has been reached.
            if ($position > $currentRowOffset && $position == $nextRowOffset) {
                break;
            }

            switch ($code) {
                case OLEReader::XLS_WORKSHEET_BLANK:
                    $rowIndex = $this->worksheet->getInt($position, $this->workbook, 2);
                    $columnIndex = $this->worksheet->getInt($position + 2, $this->workbook, 2);
                    $xfIndex = $this->worksheet->getInt($position + 4, $this->workbook, 2);
                    $cell = new Cell($rowIndex, $columnIndex);
                    $this->row[$columnIndex] = $cell->setValue('')->setXfIndex($xfIndex);
                    break;
                case OLEReader::XLS_WORKSHEET_MULBLANK:
                    $rowIndex = $this->worksheet->getInt($position, $this->workbook, 2);
                    $columnStartIndex = $this->worksheet->getInt($position + 2, $this->workbook, 2);
                    $columnEndIndex = $this->worksheet->getInt($position + $length - 2, $this->workbook, 2);
                    $columnCount = $columnEndIndex - $columnStartIndex + 1;
                    for ($i = 0; $i < $columnCount; $i++) {
                        $xfIndex = $this->worksheet->getInt($position + $i * 2 + 4, $this->workbook, 2);
                        $cell = new Cell($rowIndex, $columnStartIndex + $i);
                        $this->row[$columnStartIndex + $i] =  $cell->setXfIndex($xfIndex)->setValue('');
                    }
                    break;
                case OLEReader::XLS_WORKSHEET_RK:
                case OLEReader::XLS_WORKSHEET_RK2:
                    $rowIndex = $this->worksheet->getInt($position, $this->workbook, 2);
                    $columnIndex = $this->worksheet->getInt($position + 2, $this->workbook, 2);
                    $xfIndex = $this->worksheet->getInt($position + 4, $this->workbook, 2);
                    $rkNum = $this->worksheet->getInt($position + 6, $this->workbook);
                    $value = $this->worksheet->getIEEE754($rkNum);
                    $cell = new Cell($rowIndex, $columnIndex);
                    $this->row[$columnIndex] = $cell->setValue($value)->setXfIndex($xfIndex);
                    break;
                case OLEReader::XLS_WORKSHEET_MULRK:
                    $rowIndex = $this->worksheet->getInt($position, $this->workbook, 2);
                    $columnStartIndex = $this->worksheet->getInt($position + 2, $this->workbook, 2);
                    $columnEndIndex = $this->worksheet->getInt($position + $length - 2, $this->workbook, 2);
                    $columnCount = $columnEndIndex - $columnStartIndex + 1;

                    for ($i = 0; $i < $columnCount; $i++) {
                        $xfIndex = $this->worksheet->getInt($position + $i * 6 + 4, $this->workbook, 2);
                        $rkNum = $this->worksheet->getInt($position + ($i + 1) * 6, $this->workbook);
                        $value = $this->worksheet->getIEEE754($rkNum);
                        $cell = new Cell($rowIndex, $columnStartIndex + $i);
                        $this->row[$columnStartIndex + $i] = $cell->setValue($value)->setXfIndex($xfIndex);
                    }
                    break;
                case OLEReader::XLS_WORKSHEET_NUMBER:
                    $rowIndex = $this->worksheet->getInt($position, $this->workbook, 2);
                    $columnIndex = $this->worksheet->getInt($position + 2, $this->workbook, 2);
                    $xfIndex = $this->worksheet->getInt($position + 4, $this->workbook, 2);
                    $tmpValue = unpack("d", substr($this->workbook, $position + 6, 8));
                    $value = current($tmpValue);
                    $cell = new Cell($rowIndex, $columnIndex);
                    $this->row[$columnIndex] = $cell->setValue($value)->setXfIndex($xfIndex);
                    break;
                case OLEReader::XLS_WORKSHEET_BOOLERR:
                    $rowIndex = $this->worksheet->getInt($position, $this->workbook, 2);
                    $columnIndex = $this->worksheet->getInt($position + 2, $this->workbook, 2);
                    $value = $this->worksheet->getInt($position + 3, $this->workbook, 1);
                    $cell = new Cell($rowIndex, $columnIndex);
                    $this->row[$columnIndex] = $cell->setValue($value);
                    break;
                case OLEReader::XLS_WORKSHEET_STRING:
                    // $rowIndex = $this->worksheet->getInt($position, $this->workbook, 2);
                    // $columnIndex = $this->worksheet->getInt($position + 2, $this->workbook, 2);

                    if ($this->version == OLEReader::XLS_BIFF_VERSION_8) {
                        // Unicode 16 string, like an SST record
                        $tmpPosition = $position;
                        // $numChars =ord($data[$xpos]) | (ord($data[$xpos+1]) << 8);
                        $numChars = $this->worksheet->getInt($tmpPosition, $this->workbook, 2);
                        $tmpPosition += 2;
                        $optionFlags = $this->worksheet->getInt($tmpPosition, $this->workbook, 1);
                        $tmpPosition++;
                        $asciiEncoding = ($optionFlags & 0x01) == 0;
                        $extendedString = ($optionFlags & 0x04) != 0;
                        // See if string contains formatting information
                        $richString = (($optionFlags & 0x08) != 0);
                        if ($richString) {
                            // Read in the crun
                            $formattingRuns = $this->worksheet->getInt($tmpPosition, $this->workbook, 2);
                            $tmpPosition += 2;
                        }
                        if ($extendedString) {
                            // Read in cchExtRst
                            $extendedRunLength = $this->worksheet->getInt($tmpPosition, $this->workbook);
                            $tmpPosition += 4;
                        }
                        $len = ($asciiEncoding) ? $numChars : $numChars * 2;
                        $value = substr($this->workbook, $tmpPosition, $len);
                        $tmpPosition += $len;
                        $value = $asciiEncoding ? $value : $this->worksheet->convertEncoding($value);
                    } elseif ($this->version == OLEReader::XLS_BIFF_VERSION_7) {
                        // Simple byte string
                        $tmpPosition = $position;
                        $numChars = $this->worksheet->getInt($tmpPosition, $this->workbook, 2);
                        $tmpPosition += 2;
                        $value = substr($this->workbook, $tmpPosition, $numChars);
                    }

                    if (isset($previousRowIndex) && isset($previousColIndex)) {
                        $cell = new Cell($previousRowIndex, $previousColIndex);
                        $this->row[$previousColIndex] = $cell->setValue($value);
                    }
                    break;
                case OLEReader::XLS_WORKSHEET_RSTRING:
                    $rowIndex = $this->worksheet->getInt($position, $this->workbook, 2);
                    $columnIndex = $this->worksheet->getInt($position + 2, $this->workbook, 2);
                    $xfIndex = $this->worksheet->getInt($position + 4, $this->workbook, 2);
                    $cell = new Cell($rowIndex, $columnIndex);
                    $this->row[$columnIndex] = $cell;
                    break;
                case OLEReader::XLS_WORKSHEET_LABEL:
                    $rowIndex = $this->worksheet->getInt($position, $this->workbook, 2);
                    $columnIndex = $this->worksheet->getInt($position + 2, $this->workbook, 2);
                    $xfIndex = $this->worksheet->getInt($position + 4, $this->workbook, 2);
                    $labelLength = $this->worksheet->getInt($position + 6, $this->workbook, 2);
                    $value = substr($this->workbook, $position + 8, $labelLength);
                    $cell = new Cell($rowIndex, $columnIndex);
                    $this->row[$columnIndex] = $cell->setValue($value)->setXfIndex($xfIndex);
                    break;
                case OLEReader::XLS_WORKSHEET_LABELSST:
                    $rowIndex = $this->worksheet->getInt($position, $this->workbook, 2);
                    $columnIndex = $this->worksheet->getInt($position + 2, $this->workbook, 2);
                    $xfIndex = $this->worksheet->getInt($position + 4, $this->workbook, 2);
                    $sstIndex = $this->worksheet->getInt($position + 6, $this->workbook, 4);
                    $cell = new Cell($rowIndex, $columnIndex);
                    $this->row[$columnIndex] = $cell->setXfIndex($xfIndex)->setSstIndex($sstIndex);
                    break;
                case OLEReader::XLS_WORKSHEET_FORMULA:
                    $rowIndex = $this->worksheet->getInt($position, $this->workbook, 2);
                    $columnIndex = $this->worksheet->getInt($position + 2, $this->workbook, 2);
                    $xfIndex = $this->worksheet->getInt($position + 4, $this->workbook, 2);

                    if (
                        $this->worksheet->getInt($position + 6, $this->workbook, 1) == 0
                        && $this->worksheet->getInt($position + 12, $this->workbook, 2) == 255
                    ) {
                        //String formula. Result follows in a STRING record
                        // This row/col are stored to be referenced in that record
                        // http://code.google.com/p/php-excel-reader/issues/detail?id=4
                        $previousRowIndex = $rowIndex;
                        $previousColIndex = $columnIndex;
                    } elseif (
                        $this->worksheet->getInt($position + 12, $this->workbook, 1) == 255
                        && $this->worksheet->getInt($position + 13, $this->workbook, 1) == 255
                    ) {
                        switch ($this->worksheet->getInt($position + 6, $this->workbook, 1)) {
                            case 1:
                                //Boolean formula. Result is in +2; 0=false,1=true
                                // http://code.google.com/p/php-excel-reader/issues/detail?id=4
                                $value = $this->worksheet->getInt($position + 8, $this->workbook, 1) == 1 ? 'TRUE' : 'FALSE';
                                break;
                            case 2:
                                //Error formula. Error code is in +2;
                                break;
                            case 3:
                                //Formula result is a null string.
                                $value = '';
                                break;
                        }
                    } else {
                        // result is a number, so first 14 bytes are just like a _NUMBER record
                        $tmpValue = unpack('d', substr($this->workbook, $position + 6, 8)); // It machine machine dependent
                        $value = current($tmpValue);
                    }

                    $cell = new Cell($rowIndex, $columnIndex);
                    isset($value) && $cell->setValue($value);
                    isset($xfIndex) && $cell->setXfIndex($xfIndex);
                    $this->row[$columnIndex] = $cell;

                    break;
                case OLEReader::XLS_WORKSHEET_ARRAY:
                    break;
                case OLEReader::XLS_WORKSHEET_SHAREDFMLA:
                    break;
                case OLEReader::XLS_WORKSHEET_DATATABLE:
                    break;
                case OLEReader::XLS_WORKSHEET_DATATABLE2:
                    break;
                default:
                    // echo dechex($code) . PHP_EOL;
                    break;
            }

            $this->currentOffset = $position;
            $this->currentLength = $length;
        } while ($code != OLEReader::XLS_WORKSHEET_EOF);

        return parent::current();
    }

    /**
     * Rewind iterator.
     * 
     * @return void
     */
    public function rewind(): void
    {
        $this->worksheetVersion = $this->worksheet->getInt($this->currentOffset + 4, $this->workbook, 2);
        $this->currentOffset = $this->worksheetOffset;
        $this->currentLength = $this->worksheet->getInt($this->currentOffset + 2, $this->workbook, 2);
        $this->currentOffset += 4;

        if ($this->startRow > 1) {
            // 4.7.3 Cell Block
            do {
                $position = $this->currentOffset + $this->currentLength;

                $lowCode = $this->worksheet->getInt($position, $this->workbook, 1);
                if ($lowCode == OLEReader::XLS_WORKSHEET_EOF) {
                    break;
                }

                $code = $this->worksheet->getInt($position, $this->workbook, 2);
                $length = $this->worksheet->getInt($position + 2, $this->workbook, 2);
                $position += 4;

                $currentRowOffset = $this->rowOffsets[$this->position + 1];
                $nextRowOffset = $this->rowOffsets[$this->position + 2] ?? null;
                // Determine whether next row has been reached.
                if ($position > $currentRowOffset && $position == $nextRowOffset) {
                    $this->position++;
                }

                if ($this->position >= $this->startRow - 1) {
                    break;
                }

                $this->currentOffset = $position;
                $this->currentLength = $length;
            } while ($code != OLEReader::XLS_WORKSHEET_EOF);
        }
    }
}
