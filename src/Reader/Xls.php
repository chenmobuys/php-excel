<?php

namespace Excel\Reader;

use Excel\Reader\BaseRowIterator;
use Excel\Reader\Xls\Excel;
use Excel\Reader\Xls\OLEReader;
use Excel\Reader\Xls\RowIterator;
use Excel\Shared\Coordinate;

class Xls extends BaseReader
{
    /**
     * @var \Excel\Reader\Xls\OLEReader
     */
    private $OLEReader;

    /**
     * Load worksheet.
     * 
     * @param string $filename
     * @return void
     */
    public function load(string $filename): void
    {
        $style = $this->getStyle();
        $this->OLEReader = new OLEReader($filename, $style);
        $this->worksheetNames = $this->getWorksheetNamesByOLEReader($this->OLEReader);
        $this->worksheetInfo = $this->getWorksheetInfoByOLEReader($this->OLEReader);
    }

    /**
     * Determine whether the file is readable.
     * 
     * @param string $filename 
     * @return bool
     */
    public function isReadable(string $filename): bool
    {
        return is_readable($filename) && OLEReader::isReadable($filename);
    }

    /**
     * List worksheet names.
     * 
     * @param string $filename
     * @return array
     */
    public function listWorksheetNames(string $filename): array
    {
        $OLEReader = new OLEReader($filename);
        return $this->getWorksheetNamesByOLEReader($OLEReader);
    }

    /**
     * List worksheet info.
     * 
     * @param string $filename
     * @return array
     */
    public function listWorksheetInfo(string $filename): array
    {
        $OLEReader = new OLEReader($filename);
        return $this->getWorksheetInfoByOLEReader($OLEReader);
    }

    /**
     * Get worksheet names by OLEReader.
     * 
     * @param \Excel\Reader\Xls\OLEReader
     * 
     * @return array
     */
    private function getWorksheetNamesByOLEReader(OLEReader $OLEReader): array
    {
        $worksheets = $OLEReader->getWorksheets();
        return array_column($worksheets, 'name');
    }

    /**
     * Get worksheet info by OLEReader.
     * 
     * @param \Excel\Reader\Xls\OLEReader
     * 
     * @return array
     */
    private function getWorksheetInfoByOLEReader(OLEReader $OLEReader): array
    {
        $worksheetInfo = [];

        $worksheets = $OLEReader->getWorksheets();
        $workbook = $OLEReader->getWorkbook();

        foreach ($worksheets as $worksheet) {
            $tmpInfo = [
                'worksheetName'     => $worksheet['name'],
                'lastColumnLetter'  => 'A',
                'lastColumnIndex'   => 0,
                'totalRows'         => 0,
                'totalColumns'      => 0,
            ];

            $position = $worksheet['offset'];
            $code = $OLEReader->getInt($position, $workbook, 2);
            $length = $OLEReader->getInt($position + 2, $workbook, 2);
            $version = $OLEReader->getInt($position + 4, $workbook, 2);
            $substreamType = $OLEReader->getInt($position + 6, $workbook, 2);
            $position += 4;

            do {
                $position += $length;
                $lowCode = $OLEReader->getInt($position, $workbook, 1);

                if ($lowCode == OLEReader::XLS_WORKSHEET_EOF) {
                    break;
                }

                $code = $OLEReader->getInt($position, $workbook, 2);
                $length = $OLEReader->getInt($position + 2, $workbook, 2);
                $position += 4;

                switch ($code) {
                    case OLEReader::XLS_WORKSHEET_DIMENSION:
                        // Index to first used row
                        $rowStartIndex = $OLEReader->getInt($position, $workbook, 4);
                        // Index to last used row, increased by 1
                        $rowEndIndex = $OLEReader->getInt($position + 4, $workbook, 4);
                        // Index to first used column
                        $columnStartIndex = $OLEReader->getInt($position + 8, $workbook, 2);
                        // Index to last used column, increased by 1
                        $columnEndIndex = $OLEReader->getInt($position + 10, $workbook, 2);

                        $totalRows = $rowEndIndex;
                        $totalColumns = $columnEndIndex;
                        $tmpInfo['totalRows'] = max($tmpInfo['totalRows'], $totalRows);
                        $tmpInfo['totalColumns'] = max($tmpInfo['totalColumns'], $totalColumns);
                        break;
                    case OLEReader::XLS_WORKSHEET_EOF:
                        break 2;
                }
            } while ($code != OLEReader::XLS_WORKSHEET_EOF);

            $tmpInfo['lastColumnIndex'] = $tmpInfo['totalColumns'] - 1;
            $tmpInfo['lastColumnLetter'] = $tmpInfo['totalColumns'] ? Coordinate::columnLetterFromColumnIndex($tmpInfo['lastColumnIndex']) : null;
            $worksheetInfo[] = $tmpInfo;
        }

        return $worksheetInfo;
    }

    /**
     * Get row iterator.
     * 
     * @param array $worksheetInfo
     * @param int $worksheetIndex
     * @param int $startRow
     * @param int|null $endRow
     * @return \Excel\Reader\BaseRowIterator
     */
    protected function getRowIterator(array $worksheetInfo, int $worksheetIndex, int $startRow = 1, int $endRow = null): BaseRowIterator
    {
        $worksheet = $this->OLEReader;
        $style = $this->getStyle();

        return new RowIterator($worksheet, $style, $worksheetIndex, $startRow, $endRow);
    }
}
