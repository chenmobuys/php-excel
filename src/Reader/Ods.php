<?php

namespace Excel\Reader;

use Excel\Reader\Ods\RowIterator;
use Excel\Shared\Coordinate;
use XMLReader;
use ZipArchive;

class Ods extends BaseReader
{
    /**
     * @var ZipArchive $zip
     */
    private $zip;

    /**
     * @var array
     */
    private $worksheetPaths;

    /**
     * Load worksheet.
     * 
     * @param string $filename
     * @return void
     */
    public function load(string $filename): void
    {
        $this->zip = new ZipArchive();
        $this->zip->open($filename);

        $this->worksheetNames = $this->getWorksheetNamesByZip($this->zip);
        $this->worksheetInfo = $this->getWorksheetInfoByZip($this->zip);
    }

    /**
     * Determine whether the file is readable.
     * 
     * @param string $filename 
     * @return bool
     */
    public function isReadable(string $filename): bool
    {
        $zip = new ZipArchive();
        if ($zip->open($filename)) {
            $zip->close();
            return true;
        }

        return false;
    }

    /**
     * List worksheet names.
     * 
     * @param string $filename
     * @return array
     */
    public function listWorksheetNames(string $filename): array
    {
        $zip = new ZipArchive();
        $zip->open($filename);

        return $this->getWorksheetNamesByZip($zip);
    }

    /**
     * List worksheet info.
     * 
     * @param string $filename
     * @return array
     */
    public function listWorksheetInfo(string $filename): array
    {
        $zip = new ZipArchive();
        $zip->open($filename);

        return $this->getWorksheetInfoByZip($zip);
    }

    /**
     * Get worksheet names by ZipArchive.
     * 
     * @param \ZipArchive
     * 
     * @return array
     */
    private function getWorksheetNamesByZip(ZipArchive $zip): array
    {
        $worksheetNames = [];

        $content = $zip->getFromName('content.xml');
        /** @var XMLReader $xml */
        $xml = XMLReader::XML($content);
        while ($xml->read()) {
            if ($xml->name == 'table:table') {
                $worksheetNames[] = $xml->getAttribute('table:name');
                $xml->next();
            }
        }
        $xml->close();

        return $worksheetNames;
    }

    /**
     * Get worksheet info by ZipArchive.
     * 
     * @param \ZipArchive
     * 
     * @return array
     */
    private function getWorksheetInfoByZip(ZipArchive $zip): array
    {
        $worksheetInfo = [];

        $content = $zip->getFromName('content.xml');
        /** @var XMLReader $xml */
        $xml = XMLReader::XML($content);
        $worksheetIndex = 0;
        $emptyCellCount = 0;
        $tmpInfo = null;
        while ($xml->read()) {
            if ($xml->name == 'table:table' && $xml->nodeType == XMLReader::ELEMENT) {
                $worksheetName = $xml->getAttribute('table:name');
                $this->zip && $this->worksheetPaths[] = $worksheetIndex;
                $tmpInfo = [
                    'worksheetName'     => $worksheetName,
                    'lastColumnLetter'  => 'A',
                    'lastColumnIndex'   => 0,
                    'totalRows'         => 0,
                    'totalColumns'      => 0,
                ];

                $currCells = 0;

                do {
                    $xml->read();
                    if ($xml->name == 'table:table-row' && $xml->nodeType == XMLReader::ELEMENT) {
                        $rowspan = $xml->getAttribute('table:number-rows-repeated');
                        $rowspan = empty($rowspan) ? 1 : $rowspan;
                        $tmpInfo['totalRows'] += $rowspan;
                        $tmpInfo['totalColumns'] = max($tmpInfo['totalColumns'], $currCells);
                        $currCells = 0;
                        // Step into the row
                        $xml->read();
                        do {
                            $doread = true;
                            if ($xml->name == 'table:table-cell' && $xml->nodeType == XMLReader::ELEMENT) {
                                if (!$xml->isEmptyElement) {
                                    ++$currCells;
                                    $xml->next();
                                    $doread = false;
                                }
                            } elseif ($xml->name == 'table:covered-table-cell' && $xml->nodeType == XMLReader::ELEMENT) {
                                $mergeSize = $xml->getAttribute('table:number-columns-repeated');
                                $currCells += (int) $mergeSize;
                            }
                            if ($doread) {
                                $xml->read();
                            }
                        } while ($xml->name != 'table:table-row');
                    }
                } while ($xml->name != 'table:table');

                $tmpInfo['totalColumns'] = max($tmpInfo['totalColumns'], $currCells);
                $tmpInfo['totalRows'] = $tmpInfo['totalColumns'] > 0 ? $tmpInfo['totalRows'] : 0;
                $tmpInfo['lastColumnIndex'] = $tmpInfo['totalColumns'] - 1;
                $tmpInfo['lastColumnLetter'] = $tmpInfo['lastColumnIndex'] > -1 ? Coordinate::columnLetterFromColumnIndex($tmpInfo['lastColumnIndex']) : null;
                $worksheetInfo[] = $tmpInfo;
            }
        }

        $xml->close();

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
        $worksheetXml = $this->zip->getFromName('content.xml');
        /** @var XMLReader */
        $worksheetXmlReader = XMLReader::XML($worksheetXml);
        $style = $this->getStyle();

        return new RowIterator($worksheetXmlReader, $style, $worksheetIndex, $startRow, $endRow);
    }
}
