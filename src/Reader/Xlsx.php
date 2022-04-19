<?php

namespace Excel\Reader;

use ZipArchive;
use SimpleXMLElement;
use XMLReader;
use Excel\Settings;
use Excel\Shared\Coordinate;
use Excel\Shared\Style\Format;
use Excel\Reader\Xlsx\Namespaces;
use Excel\Reader\Xlsx\RowIterator;

class Xlsx extends BaseReader
{
    /**
     * @var ZipArchive
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

        $this->prepareBeforeIterator($this->zip);
    }

    /**
     * Determine whether the file is readable.
     * 
     * @param string $filename 
     * @return bool
     */
    public function isReadable(string $filename): bool
    {
        $result = false;
        $zip = new ZipArchive();

        if ($zip->open($filename) === true) {

            $rels = new SimpleXMLElement(
                $zip->getFromName('_rels/.rels'),
                Settings::getLibXmlLoaderOptions()
            );
            if (!$rels) {
                return $result;
            }

            foreach ($rels->Relationship as $rel) {
                $relType = (string) $rel->attributes()['Type'];
                $relTarget = (string) $rel->attributes()['Target'];
                if ($relType !== Namespaces::OFFICE_DOCUMENT) {
                    continue;
                }
                $basename = basename($relTarget);
                if (preg_match('/workbook.*\.xml/', $basename)) {
                    $result = true;
                    break;
                }
            }

            $zip->close();
        }

        return $result;
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

        $rels = new SimpleXMLElement(
            $zip->getFromName('_rels/.rels'),
            Settings::getLibXmlLoaderOptions(),
            false,
            Namespaces::RELATIONSHIPS
        );

        foreach ($rels->Relationship as $rel) {
            $relType = (string) $rel->attributes()['Type'];
            if ($relType !== Namespaces::OFFICE_DOCUMENT) {
                continue;
            }

            $workbookXml = new SimpleXMLElement(
                $zip->getFromName((string) $rel->attributes()['Target']),
                Settings::getLibXmlLoaderOptions()
            );
            if (!$workbookXml->sheets) {
                continue;
            }
            foreach ($workbookXml->sheets->sheet as $sheet) {
                $worksheetNames[] = (string) $sheet->attributes()['name'];
            }
        }

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

        $rels = new SimpleXMLElement(
            $zip->getFromName('_rels/.rels'),
            Settings::getLibXmlLoaderOptions(),
            false,
            Namespaces::RELATIONSHIPS
        );
        foreach ($rels->Relationship as $rel) {
            $relType = (string) $rel->attributes()['Type'];
            $relTarget = (string) $rel->attributes()['Target'];
            if ($relType != Namespaces::OFFICE_DOCUMENT) {
                continue;
            }

            $worksheets = [];
            $xmlNamespaceBase = dirname($relType);
            $xmlDirectoryBase = dirname($relTarget);
            $workbookRels = new SimpleXMLElement(
                $zip->getFromName($xmlDirectoryBase . '/_rels/' . basename($relTarget) . '.rels'),
                Settings::getLibXmlLoaderOptions()
            );
            foreach ($workbookRels->Relationship as $workbookRel) {
                if ((string) $workbookRel['Type'] !== $xmlNamespaceBase . '/worksheet') {
                    continue;
                }
                $worksheets[(string) $workbookRel->attributes()] = $xmlDirectoryBase . '/' . $workbookRel->attributes()['Target'];
            }

            $workbookXml = new SimpleXMLElement(
                $zip->getFromName((string) $rel->attributes()['Target']),
                Settings::getLibXmlLoaderOptions()
            );
            if (!$workbookXml->sheets) {
                continue;
            }
            foreach ($workbookXml->sheets->sheet as $sheet) {
                $worksheetPath = $worksheets[(string) $sheet->attributes($xmlNamespaceBase)['id']];
                $this->zip && $this->worksheetPaths[] = $worksheetPath;
                $tmpInfo = [
                    'worksheetName'     => (string) $sheet->attributes()['name'],
                    'lastColumnLetter'  => 'A',
                    'lastColumnIndex'   => 0,
                    'totalRows'         => 0,
                    'totalColumns'      => 0,
                ];

                /** @var XMLReader $xml */
                $xml = XMLReader::XML(
                    $zip->getFromName($worksheetPath),
                    Settings::getLibXmlLoaderOptions()
                );
                $xml->setParserProperty(XMLReader::DEFAULTATTRS, true);

                $totalColumns = 0;
                while ($xml->read()) {
                    if ($xml->nodeType != XMLReader::ELEMENT) {
                        continue;
                    }
                    if ($xml->namespaceURI != Namespaces::MAIN) {
                        continue;
                    }
                    if ($xml->localName == 'row') {
                        $row = $xml->getAttribute('r');
                        $tmpInfo['totalRows'] = $row;
                        $tmpInfo['totalColumns'] = max($tmpInfo['totalColumns'], $totalColumns);
                        $totalColumns = 0;
                    }
                    if ($xml->localName == 'c') {
                        $cell = $xml->getAttribute('r');
                        $letter = preg_replace('/[^[:alpha:]]/S', '', $cell);
                        $totalColumns = $cell ? max($totalColumns, Coordinate::columnIndexFromColumnLetter($letter) + 1) : $totalColumns;
                    }
                }
                $xml->close();

                $tmpInfo['totalColumns'] = max($tmpInfo['totalColumns'], $totalColumns);
                $tmpInfo['lastColumnIndex'] = $tmpInfo['totalColumns'] - 1;
                $tmpInfo['lastColumnLetter'] = $tmpInfo['lastColumnIndex'] > -1 ? Coordinate::columnLetterFromColumnIndex($tmpInfo['lastColumnIndex']) : null;

                $worksheetInfo[] = $tmpInfo;
            }
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
        $worksheetXml = $this->zip->getFromName($this->worksheetPaths[$worksheetIndex]);
        /** @var XMLReader $worksheetXmlReader */
        $worksheetXmlReader = XMLReader::XML($worksheetXml);
        $style = $this->getStyle();

        return new RowIterator($worksheetXmlReader, $style, $worksheetIndex, $startRow, $endRow);
    }

    /**
     * Prepare for read.
     * 
     * @param \ZipArchive $zip
     */
    private function prepareBeforeIterator(ZipArchive $zip): void
    {
        $rels = new SimpleXMLElement(
            $zip->getFromName('_rels/.rels'),
            Settings::getLibXmlLoaderOptions()
        );

        foreach ($rels->Relationship as $rel) {
            $relType = (string) $rel->attributes()['Type'];
            $relTarget = (string) $rel->attributes()['Target'];
            if ($relType !== Namespaces::OFFICE_DOCUMENT) {
                continue;
            }

            $xmlDirectoryBase = dirname($relTarget);
            $workbookRels = new SimpleXMLElement(
                $zip->getFromName($xmlDirectoryBase . '/_rels/' . basename($relTarget) . '.rels'),
                Settings::getLibXmlLoaderOptions()
            );
            $stylesAbsoluteTarget = '';
            $sharedStringsAbsoluteTarget = '';
            foreach ($workbookRels->Relationship as $workbookRel) {
                if ((string) $workbookRel->attributes()['Type'] == Namespaces::STYLES) {
                    $stylesAbsoluteTarget = (string) $workbookRel->attributes()['Target'];
                }
                if ((string) $workbookRel->attributes()['Type'] == Namespaces::SHARED_STRINGS) {
                    $sharedStringsAbsoluteTarget = (string) $workbookRel->attributes()['Target'];
                }
            }
        }

        $stylesTarget = !$stylesAbsoluteTarget ?: $xmlDirectoryBase . '/' . $stylesAbsoluteTarget;
        $sharedStringsTarget = !$sharedStringsAbsoluteTarget ?: $xmlDirectoryBase . '/' . $sharedStringsAbsoluteTarget;

        $this->prepareStyles($this->zip->getFromName($stylesTarget));
        $this->prepareSharedStrings($this->zip->getFromName($sharedStringsTarget));
        $this->prepareCalendar($workbookRels);
    }

    /**
     * Prepare styles
     * 
     * @param string $stylesXml
     * @return void
     */
    private function prepareStyles(string $stylesXml): void
    {
        if (!$stylesXml) {
            return;
        }

        $styles = new SimpleXMLElement($stylesXml);
        if (!($styles && $styles->cellXfs && $styles->cellXfs->xf)) {
            return;
        }

        foreach ($styles->cellXfs->xf as $xf) {
            // Format #0 is a special case - it is the "General" format that is applied regardless of applyNumberFormat
            if ($xf->attributes()->applyNumberFormat || (0 == (int) $xf->attributes()->numFmtId)) {
                // If format ID >= 164, it is a custom format and should be read from styleSheet\numFmts
                $this->getStyle()->getXF()->append((int) $xf->attributes()->numFmtId);
            } else {
                // 0 for "General" format
                $this->getStyle()->getXF()->append(0);
            }
        }

        // foreach ($styles->cellStyleXfs->xf as $xf) {
        // }

        if ($styles->numFmts && $styles->numFmts->numFmt) {
            foreach ($styles->numFmts->numFmt as $numFmt) {
                $this->getStyle()->getFormat()->append((int) $numFmt->attributes()->numFmtId, (string) $numFmt->attributes()->formatCode);
            }
        }

        unset($styles);
    }

    /**
     * Prepare shared strings.
     * 
     * @param string $sharedStringsXml
     * @return void
     */
    private function prepareSharedStrings($sharedStringsXml): void
    {
        if (!$sharedStringsXml) {
            return;
        }

        /** @var XMLReader */
        $sharedStrings = XMLReader::XML($sharedStringsXml);
        $sharedStringsCount = 0;

        while ($sharedStrings->read()) {
            if ($sharedStrings->name == 'sst') {
                $sharedStringsCount = $sharedStrings->getAttribute('count');
                if (is_null($sharedStringsCount)) {
                    $sharedStringsCount = $sharedStrings->getAttribute('uniqueCount');
                }
                break;
            }
        }

        if (!$sharedStringsCount) {
            return;
        }

        $cacheIndex = 0;
        $cacheValue = '';
        while ($sharedStrings->read()) {
            switch ($sharedStrings->name) {
                case 'si':
                    if ($sharedStrings->nodeType == XMLReader::END_ELEMENT) {
                        $this->getStyle()->getSST()->append($cacheIndex, $cacheValue);
                        $cacheIndex++;
                        $cacheValue = '';
                    }
                    break;
                case 't':
                    if ($sharedStrings->nodeType == XMLReader::END_ELEMENT) {
                        break;
                    }
                    $cacheValue .= $sharedStrings->readString();
                    break;
            }
        }

        $sharedStrings->close();
    }

    /**
     * Prepare calendar.
     * 
     * @param \SimpleXMLElement $workbookRels
     * @return void
     */
    private function prepareCalendar($workbookRels): void
    {
        if ($workbookRels->workbookPr) {
            $attrs1904 = (string)  $workbookRels->workbookPr->attributes()['date1904'] ?? false;
            if ($attrs1904) {
                $this->getStyle()->getFormat()->setCalendar(Format::CALENDAR_MAC_1904);
            }
        }
    }

    // private function prepareHyperlinks()
    // {
    //     foreach ($this->worksheetInfo as $worksheetInfo) {
    //         $worksheetRelsPath = dirname($worksheetInfo['worksheetPath']) . '/_rels/' . basename($worksheetInfo['worksheetPath']) . '.rels';
    //         if ($this->zip->locateName($worksheetRelsPath)) {
    //             $worksheetRels = new SimpleXMLElement(
    //                 $this->zip->getFromName($worksheetRelsPath),
    //                 Settings::getLibXmlLoaderOptions()
    //             );
    //             foreach ($worksheetRels->Relationship as $worksheetRel) {
    //                 if ((string) $worksheetRel->attributes()['Type'] != Namespaces::HYPERLINK) {
    //                     continue;
    //                 }

    //                 $hyperlinksAbsoluteTarget = (string) $worksheetRel->attributes()['Target'];
    //                 $this->hyperlinks[$worksheetInfo['worksheetPath']][(string) $worksheetRel->attributes()['Id']] = $hyperlinksAbsoluteTarget;
    //             }
    //         }

    //         if (!isset($this->hyperlinks[$worksheetInfo['worksheetPath']])) {
    //             continue;
    //         }

    //         $worksheet = new SimpleXMLElement(
    //             $this->zip->getFromName($worksheetInfo['worksheetPath']),
    //             Settings::getLibXmlLoaderOptions()
    //         );

    //     }
    // }

    public function __destruct()
    {
        $this->zip && $this->zip->close();
        $this->style = null;
    }
}
