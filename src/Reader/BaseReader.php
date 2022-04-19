<?php

namespace Excel\Reader;

use Excel\Exception\SpreadsheetException;
use Excel\Shared\Format;
use Excel\Shared\Style;

/**
 * Base reader class.
 * 
 * @package \Excel\Reader
 */
abstract class BaseReader
{

    /**
     * Worksheet style
     * 
     * @var \Excel\Shared\Style
     */
    protected $style;

    /**
     * Worksheet names.
     * 
     * @var array $worksheetNames
     */
    protected $worksheetNames = [];

    /**
     * Worksheet info.
     * 
     * @var array $worksheetInfo
     */
    protected $worksheetInfo = [];

    /**
     * Load worksheet.
     * 
     * @param string $filename
     * @return void
     */
    abstract public function load(string $filename): void;

    /**
     * Determine whether the file is readable.
     * 
     * @param string $filename 
     * @return bool
     */
    abstract public function isReadable(string $filename): bool;

    /**
     * List worksheet names.
     * 
     * @param string $filename
     * @return array
     * @see \Excel\Reader\BaseReader::getWorksheetNames()
     */
    abstract public function listWorksheetNames(string $filename): array;

    /**
     * List worksheet info.
     * 
     * @param string $filename
     * @return array
     * @see \Excel\Reader\BaseReader::getWorksheetInfo()
     */
    abstract public function listWorksheetInfo(string $filename): array;

    /**
     * Get row iterator.
     * 
     * @param string $worksheetName
     * @param int $startRow
     * @param int|null $endRow
     * @return \Excel\Reader\BaseRowIterator
     */
    abstract protected function getRowIterator(array $worksheetInfo, int $worksheetIndex, int $startRow = 1, int $endRow = null): BaseRowIterator;

    /**
     * Get row iterator by worksheet name.
     * 
     * @param string $worksheetName
     * @param int $startRow
     * @param int|null $endRow
     * @return \Excel\Reader\BaseRowIterator
     */
    public function getRowIteratorByWorksheetName(string $worksheetName, int $startRow = 1, int $endRow = null): BaseRowIterator
    {
        $worksheetInfo = array_filter(
            $this->worksheetInfo,
            function ($worksheetInfo) use ($worksheetName) {
                return $worksheetInfo['worksheetName'] == $worksheetName;
            }
        );

        $worksheetIndex = key($worksheetInfo);
        $worksheetInfo = current($worksheetInfo);

        if (!$worksheetInfo) {
            throw new SpreadsheetException('Worksheet name "' . $worksheetName . '" not exists.');
        }

        $endRow = $endRow ?: $worksheetInfo['totalRows'];

        return $this->getRowIterator($worksheetInfo, $worksheetIndex, $startRow, $endRow);
    }

    /**
     * Get row iterator by worksheet index.
     * 
     * @param int $worksheetIndex
     * @param int $startRow
     * @param int|null $endRow
     * @return \Excel\Reader\BaseRowIterator
     */
    public function getRowIteratorByWorksheetIndex(int $worksheetIndex, int $startRow = 1, int $endRow = null): BaseRowIterator
    {
        $worksheetInfo = $this->worksheetInfo[$worksheetIndex] ?? null;

        if (!$worksheetInfo) {
            throw new SpreadsheetException('Worksheet index "' . $worksheetIndex . '" not exists.');
        }

        $endRow = $endRow ?: $worksheetInfo['totalRows'];

        return $this->getRowIterator($worksheetInfo, $worksheetIndex, $startRow, $endRow);
    }

    /**
     * Get acitve row iterator by worksheet name.
     * 
     * @param int $startRow
     * @param int $endRow
     * @return \Excel\Reader\BaseRowIterator
     */
    public function getActiveRowIterator(int $startRow = 1, int $endRow = null): BaseRowIterator
    {
        $worksheetIndex = key($this->worksheetNames);
        return $this->getRowIteratorByWorksheetIndex($worksheetIndex, $startRow, $endRow);
    }

    /**
     * Get all row iterators by worksheet name.
     * 
     * @return \Excel\Reader\BaseRowIterator[]
     */
    public function getRowIterators(): array
    {
        return array_map(function ($worksheetName) {
            return $this->getRowIteratorByWorksheetName($worksheetName);
        }, $this->getWorksheetNames());
    }

    /**
     * Get worksheet names.
     * 
     * @return array
     * 
     * @example
     * [
     *     'Sheet1',
     *     'Sheet2',
     *     'Sheet3',
     * ]
     */
    public function getWorksheetNames(): array
    {
        return $this->worksheetNames;
    }

    /**
     * Get worksheet info.
     * 
     * @return array
     * 
     * @example
     * [
     *     [
     *          'worksheetName' => 'Sheet1',
     *          'lastColumnLetter' => 'C', 
     *          'lastColumnIndex' => '2', 
     *          'totalRows' => '2', 
     *          'totalColumns' => '3', 
     *     ],
     *     [
     *          'worksheetName' => 'Sheet2',
     *          'lastColumnLetter' => 'C', 
     *          'lastColumnIndex' => '2', 
     *          'totalRows' => '2', 
     *          'totalColumns' => '3', 
     *     ],
     *     [
     *          'worksheetName' => 'Sheet3',
     *          'lastColumnLetter' => 'C', 
     *          'lastColumnIndex' => '2', 
     *          'totalRows' => '2', 
     *          'totalColumns' => '3', 
     *     ],
     * ]
     */
    public function getWorksheetInfo(): array
    {
        return $this->worksheetInfo;
    }

    /**
     * Get worksheet style.
     * 
     * @return \Excel\Shared\Style
     */
    public function getStyle(): Style
    {
        if (!$this->style) {
            $this->style = new Style();
        }
        return $this->style;
    }
}
