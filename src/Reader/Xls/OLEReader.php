<?php

namespace Excel\Reader\Xls;

use Excel\Exception\SpreadsheetException;
use Excel\Shared\CodePage;
use Excel\Shared\Style;
use Excel\Shared\Style\Format;

/**
 * @see http://www.openoffice.org/sc/excelfileformat.pdf
 * @see http://www.openoffice.org/sc/compdocfileformat.pdf
 * @see https://en.wikipedia.org/wiki/Microsoft_Excel
 */
class OLEReader
{
    // Stream constant
    const XLS_SECTOR_SIZE_POSITION = 0x1e;
    const XLS_SHORT_SECTOR_SIZE_POSITION = 0x20;
    const XLS_SECTOR_COUNT_POSITION = 0x2c;
    const XLS_DIRECTORY_FIRST_POSITION = 0x30;
    const XLS_STANDARD_STREAM_MIN_SIZE_POSITION = 0x38;
    const XLS_SHORT_SECTOR_FIRST_POSITION = 0x3c;
    const XLS_SHORT_SECTOR_COUNT_POSITION = 0x40;
    const XLS_MASTER_SECTOR_FIRST_POSITION = 0x44;
    const XLS_MASTER_SECTOR_COUNT_POSITION = 0x48;
    const XLS_MASTER_SECTOR_POSITION = 0x4c;
    const XLS_DIRECTORY_PROPERTY_LENGTH = 0x80;
    const XLS_DIRECTORY_PROPERTY_NAME_SIZE_POSITION = 0x40;
    const XLS_DIRECTORY_PROPERTY_TYPE_POSITION = 0x42;
    const XLS_DIRECTORY_PROPERTY_FIRST_POSITION = 0x74;
    const XLS_DIRECTORY_PROPERTY_SIZE = 0x78;

    // Worksheet constant
    const XLS_WORKSHEET_FORMULA = 0x06;
    const XLS_WORKSHEET_FORMULA2 = 0x406;
    const XLS_WORKSHEET_EOF = 0x0a;
    const XLS_WORKSHEET_CALCCOUNT = 0x0c;
    const XLS_WORKSHEET_CALCMODE = 0x0d;
    const XLS_WORKSHEET_PRECISION = 0x0e;
    const XLS_WORKSHEET_REFMODE = 0x0f;
    const XLS_WORKSHEET_DELTA = 0x10;
    const XLS_WORKSHEET_ITERATION = 0x11;
    const XLS_WORKSHEET_PROTECT = 0x12;
    const XLS_WORKSHEET_PASSWORD = 0x13;
    const XLS_WORKSHEET_HEADER = 0x14;
    const XLS_WORKSHEET_FOOTER = 0x15;
    const XLS_WORKSHEET_EXTERNSHEET = 0x17;
    const XLS_WORKSHEET_DEFINEDNAME = 0x18;
    const XLS_WORKSHEET_WINDOWPROTECT = 0x19;
    const XLS_WORKSHEET_VERTICALPAGEBREAKS = 0x1a;
    const XLS_WORKSHEET_HORIZONTALPAGEBREAKS = 0x1b;
    const XLS_WORKSHEET_NOTE = 0x1c;
    const XLS_WORKSHEET_SELECTION = 0x1d;
    const XLS_WORKSHEET_DATEMODE = 0x22;
    const XLS_WORKSHEET_EXTERNALNAME = 0x23;
    const XLS_WORKSHEET_LEFTMARGIN = 0x26;
    const XLS_WORKSHEET_RIGHTMARGIN = 0x27;
    const XLS_WORKSHEET_TOPMARGIN = 0x28;
    const XLS_WORKSHEET_BOTTOMMARGIN = 0x29;
    const XLS_WORKSHEET_PRINTHEADERS = 0x2a;
    const XLS_WORKSHEET_PRINTGRIDLINES = 0x2b;
    const XLS_WORKSHEET_FILEPASS = 0x2f;
    const XLS_WORKSHEET_FONT = 0x31;
    const XLS_WORKSHEET_CONTINUE = 0x3c;
    const XLS_WORKSHEET_WINDOW1 = 0x3d;
    const XLS_WORKSHEET_BACKUP = 0x40;
    const XLS_WORKSHEET_PANE = 0x41;
    const XLS_WORKSHEET_CODEPAGE = 0x42;
    const XLS_WORKSHEET_PLS = 0x4d;
    const XLS_WORKSHEET_DCONREF = 0x51;
    const XLS_WORKSHEET_DEFCOLWIDTH = 0x55;
    const XLS_WORKSHEET_XCT = 0x59;
    const XLS_WORKSHEET_CRN = 0x5a;
    const XLS_WORKSHEET_FILESHARING = 0x5b;
    const XLS_WORKSHEET_WRITEACCESS = 0x5c;
    const XLS_WORKSHEET_UNCALCED = 0x5e;
    const XLS_WORKSHEET_SAVERECALC = 0x5f;
    const XLS_WORKSHEET_OBJECTPROTECT = 0x63;
    const XLS_WORKSHEET_COLINFO = 0x7d;
    const XLS_WORKSHEET_GUTS = 0x80;
    const XLS_WORKSHEET_SHEETPR = 0x81;
    const XLS_WORKSHEET_GRIDSET = 0x82;
    const XLS_WORKSHEET_HCENTER = 0x83;
    const XLS_WORKSHEET_VCENTER = 0x84;
    const XLS_WORKSHEET_SHEET = 0x85;
    const XLS_WORKSHEET_WRITEPROT = 0x86;
    const XLS_WORKSHEET_COUNTRY = 0x8c;
    const XLS_WORKSHEET_HIDEOBJ = 0x8d;
    const XLS_WORKSHEET_SORT = 0x90;
    const XLS_WORKSHEET_PALETTE = 0x92;
    const XLS_WORKSHEET_STANDARDWIDTH = 0x99;
    const XLS_WORKSHEET_SCL = 0xa0;
    const XLS_WORKSHEET_PAGESETUP = 0xa1;
    const XLS_WORKSHEET_MULRK = 0xbd;
    const XLS_WORKSHEET_MULBLANK = 0xbe;
    const XLS_WORKSHEET_RSTRING = 0xd6;
    const XLS_WORKSHEET_DBCELL = 0xd7;
    const XLS_WORKSHEET_BOOKBOOL = 0xda;
    const XLS_WORKSHEET_SCENPROTECT = 0xdd;
    const XLS_WORKSHEET_XF = 0xe0;
    const XLS_WORKSHEET_MERGEDCELLS = 0xe5;
    const XLS_WORKSHEET_BITMAP = 0xe9;
    const XLS_WORKSHEET_PHONETICPR = 0xef;
    const XLS_WORKSHEET_SST = 0xfc;
    const XLS_WORKSHEET_LABELSST = 0xfd;
    const XLS_WORKSHEET_EXTSST = 0xff;
    const XLS_WORKSHEET_LABELRANGES = 0x15f;
    const XLS_WORKSHEET_USESELFS = 0x160;
    const XLS_WORKSHEET_DSF = 0x161;
    const XLS_WORKSHEET_EXTERNALBOOK = 0x1ae;
    const XLS_WORKSHEET_CFHEADER = 0x1b0;
    const XLS_WORKSHEET_DATAVALIDATIONS = 0x1b2;
    const XLS_WORKSHEET_HYPERLINK = 0x1b8;
    const XLS_WORKSHEET_DATAVALIDATION = 0x1be;
    const XLS_WORKSHEET_DIMENSION = 0x200;
    const XLS_WORKSHEET_BLANK = 0x201;
    const XLS_WORKSHEET_NUMBER = 0x203;
    const XLS_WORKSHEET_LABEL = 0x204;
    const XLS_WORKSHEET_BOOLERR = 0x205;
    const XLS_WORKSHEET_STRING = 0x207;
    const XLS_WORKSHEET_ROW = 0x208;
    const XLS_WORKSHEET_INDEX = 0x20b;
    const XLS_WORKSHEET_ARRAY = 0x221;
    const XLS_WORKSHEET_DEFAULTROWHEIGHT = 0x225;
    const XLS_WORKSHEET_DATATABLE = 0x236;
    const XLS_WORKSHEET_DATATABLE2 = 0x37;
    const XLS_WORKSHEET_WINDOW2 = 0x23e;
    const XLS_WORKSHEET_RK = 0x7e;
    const XLS_WORKSHEET_RK2 = 0x27e;
    const XLS_WORKSHEET_STYLE = 0x293;
    const XLS_WORKSHEET_FORMAT = 0x41e;
    const XLS_WORKSHEET_SHAREDFMLA = 0x4bc;
    const XLS_WORKSHEET_QUICKTIP = 0x800;
    const XLS_WORKSHEET_BOF = 0x809;
    const XLS_WORKSHEET_SHEETLAYOUT = 0x862;
    const XLS_WORKSHEET_SHEETPROTECTION = 0x867;
    const XLS_WORKSHEET_RANGEPROTECTION = 0x868;

    // Binary interchange file format versions (BIFF)
    const XLS_BIFF_VERSION_8 = 0x600;
    const XLS_BIFF_VERSION_7 = 0x500;

    // Stream types
    const XLS_SUBSTREAM_WORKBOOK_GLOBALS = 0x5;

    /**
     * Sector size
     * 
     * @var int 
     */
    private $sectorSize;

    /**
     * Sector chains
     * 
     * @var array
     */
    private $sectorChains = [];

    /**
     * Short sector size
     * 
     * @var int
     */
    private $shortSectorSize;

    /**
     * Short sector chains
     * 
     * @var array
     */
    private $shortSectorChains = [];

    /**
     * Directory stream properties
     * 
     * @var array
     */
    private $directoryProperties = [];

    /**
     * Standard stream min size
     * 
     * @var int
     */
    private $standardStreamMinSize;

    /**
     * Spreadsheet stream
     * 
     * @var string
     */
    private $spreadsheet;

    /**
     * Workbook stream
     * 
     * @var string
     */
    private $workbook;

    /**
     * Code page
     * 
     * @var int
     */
    private $codePage;

    /**
     * Date mode
     * 
     * @var int
     */
    private $dateMode;

    /**
     * Worksheet info
     * 
     * @var array
     */
    private $worksheets = [];

    /**
     * Style
     * 
     * @var \Excel\Shared\Style
     */
    private $style;

    /**
     * All row's offset
     * 
     * @var array
     */
    private $rowOffsets = [];

    /**
     * OLEReader constructor.
     * 
     * @param string $filename
     * @param \Excel\Shared\Style $style
     */
    public function __construct(string $filename, Style $style = null)
    {
        $this->spreadsheet = file_get_contents($filename);

        $this->style = $style;

        $this->readSectorChains();

        $this->readDirectoryStream();

        $this->readWorkbookStream();
    }

    /**
     * Get workbook stream.
     * 
     * @return string
     */
    public function getWorkbook(): string
    {
        return $this->workbook;
    }

    /**
     * Get worksheets' info.
     * 
     * @return array
     */
    public function getWorksheets(): array
    {
        return $this->worksheets;
    }

    /**
     * Get worksheet's info by worksheet index.
     * 
     * @return array
     */
    public function getWorksheetByIndex($worksheetIndex): array
    {
        return $this->worksheets[$worksheetIndex];
    }

    /**
     * Get row's offset by worksheet index.
     * 
     * @return array
     */
    public function getRowOffsetsByIndex($worksheetIndex): array
    {
        foreach ($this->worksheets as $index => $worksheet) {
            $tmpOffsets = [];
            // $tmpMergedCells = [];
            // $tmpHyperlinks = [];

            if ($worksheetIndex !== $index) {
                continue;
            }

            if (isset($this->rowOffsets[$worksheetIndex])) {
                return $this->rowOffsets[$worksheetIndex];
            }
            $this->rowOffsets[$index] = [];

            $position = $worksheet['offset'];
            $code = $this->getInt($position, $this->workbook, 2);
            $length = $this->getInt($position + 2, $this->workbook, 2);
            $version = $this->getInt($position + 4, $this->workbook, 2);
            $substreamType = $this->getInt($position + 6, $this->workbook, 2);
            $position += 4;

            do {
                $position += $length;
                $lowCode = $this->getInt($position, $this->workbook, 1);

                if ($lowCode == OLEReader::XLS_WORKSHEET_EOF) {
                    break;
                }

                $code = $this->getInt($position, $this->workbook, 2);
                $length = $this->getInt($position + 2, $this->workbook, 2);
                $position += 4;

                switch ($code) {
                    case OLEReader::XLS_WORKSHEET_ROW:
                        $rowIndex = $this->getInt($position, $this->workbook, 2);
                        $tmpOffsets[$rowIndex] = 0;
                        break;
                    case OLEReader::XLS_WORKSHEET_DBCELL:
                        $relativeOffset = $this->getInt($position, $this->workbook, 4);
                        $firstRowOffset = $position - $relativeOffset + 20;

                        $sumOffset = 0;
                        foreach ($tmpOffsets as $rowIndex => $tmpOffset) {
                            $sumOffset += $this->getInt($position + ($rowIndex % 32) * 2 + 4, $this->workbook, 2);
                            $this->rowOffsets[$index][$rowIndex + 1] = $firstRowOffset + $sumOffset;
                        }

                        $tmpOffsets = [];
                        break;
                    case OLEReader::XLS_WORKSHEET_MERGEDCELLS:
                        // $cellRanges = $this->getInt($position, $this->workbook, 2);
                        // for ($i = 0; $i < $cellRanges; $i++) {
                        //     $fr = $this->getInt($position + 8 * $i + 2, $this->workbook, 2);
                        //     $lr = $this->getInt($position + 8 * $i + 4, $this->workbook, 2);
                        //     $fc = $this->getInt($position + 8 * $i + 6, $this->workbook, 2);
                        //     $lc = $this->getInt($position + 8 * $i + 8, $this->workbook, 2);

                        //     if ($lr - $fr > 0) {
                        //         $tmpMergedCells[$fr + 1][$fc + 1]['rowspan'] = $lr - $fr + 1;
                        //     }
                        //     if ($lc - $fc > 0) {
                        //         $tmpMergedCells[$fr + 1][$fc + 1]['rowspan'] = $lc - $fc + 1;
                        //     }
                        // }
                        break;
                    case OLEReader::XLS_WORKSHEET_HYPERLINK:
                        // $rowStartIndex = $this->getInt($position, $this->workbook, 2);
                        // $rowEndIndex = $this->getInt($position + 2, $this->workbook, 2);
                        // $columnStartIndex = $this->getInt($position + 4, $this->workbook, 2);
                        // $columnEndIndex = $this->getInt($position + 6, $this->workbook, 2);

                        // for ($rowIndex = $rowStartIndex; $rowIndex <= $rowEndIndex; $rowIndex++) {
                        //     for ($columnIndex = $columnStartIndex; $columnIndex <= $columnEndIndex; $columnIndex++) {
                        //         $coordinate = Coordinate::columnLetterFromColumnIndex($columnIndex) . $rowIndex;
                        //         $tmpHyperlinks[$coordinate] = '';
                        //     }
                        // }
                        break;
                    case OLEReader::XLS_WORKSHEET_EOF:
                        break 2;
                }
            } while ($code != OLEReader::XLS_WORKSHEET_EOF);
        }

        return $this->rowOffsets[$worksheetIndex];
    }

    /**
     * Determine whether file is readable.
     * 
     * @return bool
     */
    public static function isReadable($filename): bool
    {
        return @file_get_contents($filename, false, null, 0, 8)
            == pack('CCCCCCCC', 0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1);
    }

    /**
     * Read sector chains.
     * 
     * @return void
     */
    private function readSectorChains(): void
    {
        $sectorSizeExponent = $this->getInt(self::XLS_SECTOR_SIZE_POSITION, $this->spreadsheet, 2);
        $sectorSize = pow(2, $sectorSizeExponent);

        $shortSectorSizeExponent = $this->getInt(self::XLS_SHORT_SECTOR_SIZE_POSITION, $this->spreadsheet, 2);
        $shortSectorSize = pow(2, $shortSectorSizeExponent);

        $sectorCountOrigin = $this->getInt(self::XLS_SECTOR_COUNT_POSITION, $this->spreadsheet);

        $standardStreamMinSize = $this->getInt(self::XLS_STANDARD_STREAM_MIN_SIZE_POSITION, $this->spreadsheet);

        $shortSectorFirst = $this->getInt(self::XLS_SHORT_SECTOR_FIRST_POSITION, $this->spreadsheet);

        // $shortSectorCount = $this->getInt(self::XLS_SHORT_SECTOR_COUNT_POSITION, $this->spreadsheet);

        $masterSectorFirst = $this->getInt(self::XLS_MASTER_SECTOR_FIRST_POSITION, $this->spreadsheet);

        $masterSectorCount = $this->getInt(self::XLS_MASTER_SECTOR_COUNT_POSITION, $this->spreadsheet);

        $sectorCount = $sectorCountOrigin;
        if ($masterSectorCount != 0) {
            $sectorCount = ($sectorSize - self::XLS_MASTER_SECTOR_POSITION) / 4;
        }

        $sectorSecIds = [];
        $position = self::XLS_MASTER_SECTOR_POSITION;
        for ($i = 0; $i < $sectorCount; $i++) {
            $sectorSecIds[$i] = $this->getInt($position, $this->spreadsheet);
            $position += 4;
        }

        for ($j = 0; $j < $masterSectorCount; $j++) {
            $position = ($masterSectorFirst + 1) * $sectorSize;
            $blocksToRead = min($sectorCountOrigin - $sectorCount, $sectorSize / 4 - 1);

            for ($i = $sectorCount; $i < $sectorCount + $blocksToRead; $i++) {
                $sectorSecIds[$i] = $this->getInt($position, $this->spreadsheet);
                $position += 4;
            }

            $sectorCount += $blocksToRead;
            if ($sectorCount < $sectorCountOrigin) {
                $masterSectorFirst = $this->getInt($position, $this->spreadsheet);
            }
        }

        // Read sector chains
        $position = 0;
        $sectorChains = [];
        for ($i = 0; $i < $sectorCount; $i++) {
            $position = ($sectorSecIds[$i] + 1) * $sectorSize;
            // echo "pos = $pos" . PHP_EOL;
            for ($j = 0; $j < $sectorSize / 4; $j++) {
                $sectorChains[] = $this->getInt($position, $this->spreadsheet);
                $position += 4;
            }
        }

        // Read short-sector chains
        $position = 0;
        $shortSectorChains = [];
        $shortSectorFirstSecId = $shortSectorFirst;
        while ($shortSectorFirstSecId != -2) {
            $position = ($shortSectorFirstSecId + 1) * $sectorSize;
            for ($j = 0; $j < $sectorSize / 4; $j++) {
                $shortSectorChains[] = $this->getInt($position, $this->spreadsheet);
                $position += 4;
            }
            $shortSectorFirstSecId = $sectorChains[$shortSectorFirstSecId];
        }

        $this->sectorSize = $sectorSize;
        $this->sectorChains = $sectorChains;
        $this->shortSectorSize = $shortSectorSize;
        $this->shortSectorChains = $shortSectorChains;
        $this->standardStreamMinSize = $standardStreamMinSize;
    }

    /**
     * Read diectory stream.
     * 
     * @return void
     */
    private function readDirectoryStream(): void
    {
        $position = 0;
        $firstSecId = $this->getInt(self::XLS_DIRECTORY_FIRST_POSITION, $this->spreadsheet);
        $directoryStreamData = $this->getStreamData($firstSecId);
        $directoryStreamLength = strlen($directoryStreamData);

        while ($position < $directoryStreamLength) {

            $propertyBinary = substr($directoryStreamData, $position, self::XLS_DIRECTORY_PROPERTY_LENGTH);
            $nameSize = $this->getInt(self::XLS_DIRECTORY_PROPERTY_NAME_SIZE_POSITION, $propertyBinary, 2);
            $type = $this->getInt(self::XLS_DIRECTORY_PROPERTY_TYPE_POSITION, $propertyBinary, 1);
            $secId = $this->getInt(self::XLS_DIRECTORY_PROPERTY_FIRST_POSITION, $propertyBinary);
            $size = $this->getInt(self::XLS_DIRECTORY_PROPERTY_SIZE, $propertyBinary);
            $name = '';
            for ($i = 0; $i < $nameSize; $i++) {
                $name .= $propertyBinary[$i];
            }
            $this->directoryProperties[] = [
                'name' => str_replace("\x00", '', $name),
                'type' => $type,
                'secId' => $secId,
                'size' => $size
            ];

            $position += self::XLS_DIRECTORY_PROPERTY_LENGTH;
        }
    }

    /**
     * Read workbook stream.
     * 
     * @return void
     */
    private function readWorkbookStream(): void
    {
        $this->workbook = $this->getWorkbookStreamData();

        $position = 0;
        $code = $this->getInt($position, $this->workbook, 2);
        $length = $this->getInt($position + 2, $this->workbook, 2);
        $version = $this->getInt($position + 4, $this->workbook, 2);
        $substreamType = $this->getInt($position + 6, $this->workbook, 2);

        $position += 4;
        do {
            $position += $length;
            $code = $this->getInt($position, $this->workbook, 2);
            $length = $this->getInt($position + 2, $this->workbook, 2);
            $position += 4;
            switch ($code) {
                case self::XLS_WORKSHEET_WINDOW1:
                case self::XLS_WORKSHEET_BACKUP:
                case self::XLS_WORKSHEET_HIDEOBJ:
                case self::XLS_WORKSHEET_BOOKBOOL:
                case self::XLS_WORKSHEET_STYLE:
                case self::XLS_WORKSHEET_PALETTE:
                case self::XLS_WORKSHEET_USESELFS:
                case self::XLS_WORKSHEET_COUNTRY:
                case self::XLS_WORKSHEET_DEFINEDNAME:
                case self::XLS_WORKSHEET_DSF:
                    // Do nothing.
                    break;
                case self::XLS_WORKSHEET_CODEPAGE:
                    $codePageNumber = $this->getInt($position, $this->workbook, 2);
                    $this->codePage = CodePage::numberToName($codePageNumber);
                    break;
                case self::XLS_WORKSHEET_DATEMODE:
                    // 0 = Base date is 1899-Dec-31 (the cell value 1 represents 1900-Jan-01)
                    // 1 = Base date is 1904-Jan-01 (the cell value 1 represents 1904-Jan-02)
                    $this->dateMode = $this->getInt($position, $this->workbook, 1);
                    $this->style && $this->dateMode && $this->style->getFormat()->setCalendar(Format::CALENDAR_MAC_1904);
                    break;
                case self::XLS_WORKSHEET_PRECISION:
                    // This record stores if formulas use the real cell values for calculation or the values displayed on the screen.
                    // 0 = Use displayed values; 1 = Use real cell values
                    $this->precision = $this->getInt($position, $this->workbook, 2);
                    break;
                case self::XLS_WORKSHEET_FONT:
                    // Font
                    break;
                case self::XLS_WORKSHEET_FORMAT:
                    $formatIndex = $this->getInt($position, $this->workbook, 2);
                    $formatCodeSize = $this->getInt($position + 2, $this->workbook, 2);
                    $formatCode = $this->readUnicodeString(substr($this->workbook, $position + 4, $formatCodeSize * 2 + 1), $formatCodeSize);
                    // Append format style
                    $this->style && $this->style->getFormat()->append($formatIndex, $formatCode);
                    break;
                case self::XLS_WORKSHEET_XF:
                    // offset:  0; size: 2; Index to FONT record
                    $fontIndex = $this->getInt($position, $this->workbook, 2);
                    $fontIndex = $fontIndex >= 4 ? ($fontIndex - 1) : $fontIndex;

                    // offset:  2; size: 2; Index to FORMAT record
                    $formatIndex = $this->getInt($position + 2, $this->workbook, 2);

                    // offset:  4; size: 2; XF type, cell protection, and parent style XF
                    // bit 2-0; mask 0x0007; XF_TYPE_PROT
                    $xfTypeProt = $this->getInt($position + 4, $this->workbook, 2);

                    // bit 0; mask 0x01; 1 = cell is locked
                    $cellLocked = (0x01 & $xfTypeProt) >> 0;

                    // bit 1; mask 0x02; 1 = Formula is hidden
                    $formulaHidden = (0x02 & $xfTypeProt) >> 1;

                    // bit 2; mask 0x04; 0 = Cell XF, 1 = Cell Style XF
                    $isCellStyleXf = (0x04 & $xfTypeProt) >> 2;

                    // offset: 6; size: 1; Alignment and text break
                    $alignmentAndTextBreakFlag = $this->getInt($position + 6, $this->workbook, 1);

                    // offset: 7; size: 1; XF_ROTATION: Text rotation angle
                    $xfRotation = $this->getInt($position + 7, $this->workbook, 1);

                    // offset: 8; size: 1; Indentation, shrink to cell size, and text direction
                    $indentation = $this->getInt($position + 8, $this->workbook, 1);

                    // offset: 9; size: 1; XF_USED_ATTRIB: Flags for used attribute groups
                    $xfUsedAttrib = $this->getInt($position + 9, $this->workbook, 1);

                    // offset: 10; size: 4; Cell border lines and background area
                    $mixedFlags1 = $this->getInt($position + 10, $this->workbook);
                    $borderLeftStyle = 0x0000000F & $mixedFlags1 >> 0;
                    $borderRightStyle = 0x000000F0 & $mixedFlags1 >> 4;
                    $borderTopStyle = 0x00000F00 & $mixedFlags1 >> 8;
                    $borderBottomStyle = 0x0000F000 & $mixedFlags1 >> 12;
                    $borderLeftColorIndex = 0x007F0000 & $mixedFlags1 >> 16;
                    $borderRightColorIndex = 0x3F800000 & $mixedFlags1 >> 23;
                    $diagonalDown = (0x40000000 & $mixedFlags1) >> 30 ? true : false;
                    $diagonalUp = (0x80000000 & $mixedFlags1) >> 31 ? true : false;

                    // offset: 14; size: 4; Cell border lines and background area
                    $mixedFlags2 = $this->getInt($position + 14, $this->workbook);
                    $borderTopColorIndex = 0x0000007F & $mixedFlags2 >> 0;
                    $borderBottomColorIndex = 0x00003F80 & $mixedFlags2 >> 7;
                    $borderDiagonalColorIndex = 0x001FC000 & $mixedFlags2 >> 14;
                    $borderDiagonalStyle = 0x01E00000 & $mixedFlags2 >> 21;
                    $fillPattern = 0xFC000000 & $mixedFlags2 >> 26;

                    // offset: 18; size: 2; Cell border lines and background area
                    $mixedFlags3 = $this->getInt(18, $this->workbook, 2);
                    $startColorIndex = 0x007F & $mixedFlags2 >> 0;
                    $endColorIndex = 0x3F80 & $mixedFlags2 >> 7;

                    // Only process formatCode
                    $this->style && $this->style->getXF()->append($formatIndex);
                    break;
                case self::XLS_WORKSHEET_SST:
                    $sPosition = $position;
                    $limitPosition = $position + $length;
                    $totalStringsCount = $this->getInt($sPosition, $this->workbook, 4);
                    $uniqueStringsCount = $this->getInt($sPosition + 4, $this->workbook, 4);
                    $sPosition += 8;
                    for ($i = $uniqueStringsCount; $i--;) {
                        // Read in the number of characters
                        if ($sPosition == $limitPosition) {
                            $opcode = $this->getInt($sPosition, $this->workbook, 2);
                            $conlength = $this->getInt(0, $this->workbook, 2);
                            if ($opcode != 0x3c) {
                                return;
                            }
                            $sPosition += 4;
                            $limitPosition = $sPosition + $conlength;
                        }
                        $stringLength = $this->getInt($sPosition, $this->workbook, 2);
                        $sPosition += 2;
                        $optionFlags = substr($this->workbook, $sPosition, 1);
                        $sPosition += 1;
                        $isCompressed = !((0x01 & ord($optionFlags)) >> 0);
                        $hasAsian = (0x04) & ord($optionFlags) >> 2;
                        $hasRichText = (0x08) & ord($optionFlags) >> 3;

                        if ($hasRichText) {
                            $richTextFormattingRunsNum = $this->getInt($sPosition, $this->workbook, 2);
                            $sPosition += 2;
                        }
                        if ($hasAsian) {
                            $asianPhoneticSettingsBlockSize = $this->getInt($sPosition, $this->workbook, 4);
                            $sPosition += 4;
                        }

                        $sLength = $isCompressed ? $stringLength : $stringLength * 2;
                        if ($sPosition + $sLength < $limitPosition) {
                            $string = substr($this->workbook, $sPosition, $sLength);
                            $sPosition += $sLength;
                        } else {
                            // found countinue
                            $bytesRead = $limitPosition - $sPosition;
                            $string = substr($this->workbook, $sPosition, $bytesRead);
                            $charsLeft = $stringLength - (($hasAsian) ? $bytesRead : ($bytesRead / 2));
                            $sPosition = $limitPosition;

                            while ($charsLeft > 0) {
                                $opcode = $this->getInt($sPosition, $this->workbook, 2);
                                $conlength = $this->getInt($sPosition + 2, $this->workbook, 2);
                                if ($opcode != 0x3c) {
                                    return;
                                }
                                $sPosition += 4;
                                $limitPosition = $sPosition + $conlength;
                                $option = $this->getInt($sPosition, $this->workbook, 1);
                                $sPosition += 1;

                                if ($hasAsian && ($option == 0)) {
                                    $sLength = min($charsLeft, $limitPosition - $sPosition);
                                    $string .= substr($this->workbook, $sPosition, $sLength);
                                    $charsLeft -= $sLength;
                                    $hasAsian = true;
                                } elseif (!$hasAsian && ($option != 0)) {
                                    $sLength = min($charsLeft * 2, $limitPosition - $sPosition);
                                    $string .= substr($this->workbook, $sPosition, $sLength);
                                    $charsLeft -= $sLength / 2;
                                    $hasAsian = false;
                                } elseif (!$hasAsian && ($option == 0)) {
                                    // Bummer - the string starts off as Unicode, but after the
                                    // continuation it is in straightforward ASCII encoding
                                    $len = min($charsLeft, $limitPosition - $sPosition);
                                    for ($j = 0; $j < $len; $j++) {
                                        $string .= substr($this->workbook, $sPosition + $j, 1) . chr(0);
                                    }
                                    $charsLeft -= $len;
                                    $hasAsian = false;
                                } else {
                                    $newstr = '';
                                    for ($j = 0; $j < strlen($string); $j++) {
                                        $newstr = $string[$j] . chr(0);
                                    }
                                    $string = $newstr;
                                    $len = min($charsLeft * 2, $limitPosition - $sPosition);
                                    $string .= substr($this->workbook, $sPosition, $sLength);
                                    $charsLeft -= $len / 2;
                                    $asciiEncoding = false;
                                }

                                $sPosition += $sLength;
                            }
                        }

                        if ($hasRichText) {
                            $sPosition += 4 * $richTextFormattingRunsNum;
                        }

                        // For extended strings, skip over the extended string data
                        if ($hasAsian) {
                            $sPosition += $asianPhoneticSettingsBlockSize;
                        }

                        // Append shared string
                        $this->style && $this->style->getSST()->append($string);
                    }
                    break;

                case self::XLS_WORKSHEET_LABELSST:
                    break;
                case self::XLS_WORKSHEET_EXTSST:
                    break;
                case self::XLS_WORKSHEET_SHEET:
                    // Absolute stream position of the BOF record of the sheet represented by this record. 
                    // This field is never encrypted in protected files.
                    $offset = $this->getInt($position, $this->workbook);

                    // 0 = Visible
                    // 1 = Hidden
                    // 2 = “Very hidden”. Can only be set programmatically, e.g. with a Visual Basic macro. It is not possible to make such a sheet visible via the user interface.
                    $state = $this->getInt($position + 4, $this->workbook, 1);
                    // 00H = Worksheet
                    // 02H = Chart
                    // 06H = Visual Basic module

                    $type = $this->getInt($position + 5, $this->workbook, 1);
                    // Unicode string, 8-bit string length.
                    // From BIFF8 on, strings are always stored using UTF-16LE text encoding. 
                    // The character array is a sequence of 16-bit values. 
                    // Additionally it is possible to use a compressed format, which omits the high bytes of all characters, if they are all zero.

                    $nameSize = $this->getInt($position + 6, $this->workbook, 1);
                    $name = $this->readUnicodeString(substr($this->workbook, $position + 7, $nameSize * 2 + 1), $nameSize);

                    $this->worksheets[] = compact('name', 'type', 'offset', 'state');

                    break;
                default:
                    // Do nothing.
                    break;
            }
        } while ($code != self::XLS_WORKSHEET_EOF);
    }

    /**
     * Get wokrbook stream.
     * 
     * @return string
     */
    private function getWorkbookStreamData(): string
    {
        $rootEntryStreamData = null;
        $rootEntryProperty = $this->getDirectoryPropertyByName('ROOT ENTRY');
        $workbookProperty = $this->getDirectoryPropertyByName('WORKBOOK');

        if (($workbookProperty['size'] < $this->standardStreamMinSize)) {
            $rootEntryStreamData = $this->getStreamData($rootEntryProperty['secId']);
            return $this->getShortStreamData($workbookProperty['secId'], $rootEntryStreamData);
        }
        return $this->getStreamData($workbookProperty['secId']);
    }

    /**
     * Get stream by sector id.
     * 
     * @return string
     */
    private function getStreamData(int $secId): string
    {
        $data = '';
        $position = 0;
        $sourceData = $this->spreadsheet;
        while ($secId != -2) {
            $position = ($secId + 1) * $this->sectorSize;
            $data .= substr($sourceData, $position, $this->sectorSize);
            $secId = $this->sectorChains[$secId];
        }
        return $data;
    }

    /**
     * Get short stream by sector id.
     * 
     * @return stream
     */
    private function getShortStreamData($secId, $sourceData): string
    {
        $data = '';
        $position = 0;
        while ($secId != -2) {
            $position = ($secId + 1) * $this->shortSectorSize;
            $data .= substr($sourceData, $position, $this->shortSectorSize);
            $secId = $this->shortSectorChains[$secId];
        }
        return $data;
    }

    /**
     * Get directory property by name.
     * 
     * @return array
     */
    private function getDirectoryPropertyByName($name): array
    {
        return current(array_filter(
            $this->directoryProperties,
            function ($directoryProperty) use ($name) {
                $name = strtolower($name);
                return strtolower($directoryProperty['name']) == $name
                    || (strtolower($directoryProperty['name']) == 'book' && $name == 'workbook');
            }
        ));
    }

    /**
     * Get int
     * 
     * @return int
     */
    public function getInt($position, $data, $length = 4): int
    {
        $position = (int) $position;

        $value = ord($data[$position]);
        for ($i = 1; $i < $length; $i++) {
            $value = $value | (ord($data[$position + $i]) << (8 * $i));
        }

        return $value >= 4294967294 ? -2 : $value;
    }

    /**
     * Get IEEE754
     * 
     * @return int
     */
    public function getIEEE754($rkNum): int
    {
        if (($rkNum & 0x02) != 0) {
            $value = $rkNum >> 2;
        } else {
            //mmp
            // I got my info on IEEE754 encoding from
            // http://research.microsoft.com/~hollasch/cgindex/coding/ieeefloat.html
            // The RK format calls for using only the most significant 30 bits of the
            // 64 bit floating point value. The other 34 bits are assumed to be 0
            // So, we use the upper 30 bits of $rkNum as follows...
            $sign = ($rkNum & 0x80000000) >> 31;
            $exp = ($rkNum & 0x7ff00000) >> 20;
            $mantissa = (0x100000 | ($rkNum & 0x000ffffc));
            $value = $mantissa / pow(2, (20 - ($exp - 1023)));
            if ($sign) {
                $value = -1 * $value;
            }
            //end of changes by mmp
        }
        if (($rkNum & 0x01) != 0) {
            $value /= 100;
        }
        return $value;
    }

    /**
     * Covert string's encoding.
     * 
     * @return string
     */
    public function convertEncoding($string, $fromEncoding = null, $toEndcoding = 'UTF-8'): string
    {
        $fromEncoding = $fromEncoding ?: $this->codePage;

        if (function_exists('mb_convert_encoding')) {
            return mb_convert_encoding($string, $toEndcoding, $fromEncoding);
        }

        if (function_exists('iconv')) {
            return iconv($fromEncoding, $toEndcoding, $string);
        }

        throw new SpreadsheetException('Convert encoding failed.');
    }

    /**
     * Read unicode string.
     * 
     * @return string
     */
    public function readUnicodeString($data, $length): string
    {
        // offset: 0: size: 1; option flags
        // bit: 0; mask: 0x01; character compression (0 = compressed 8-bit, 1 = uncompressed 16-bit)
        $isCompressed = (0x01 & ord($data[0])) >> 0;

        // bit: 2; mask: 0x04; Asian phonetic settings
        $hasAsian = (0x04) & ord($data[0]) >> 2;

        // bit: 3; mask: 0x08; Rich-Text settings
        $hasRichText = (0x08) & ord($data[0]) >> 3;

        $data = $isCompressed ? $data : $this->uncompressByteString($data);

        // offset: 1: size: var; character array
        // this offset assumes richtext and Asian phonetic settings are off which is generally wrong
        // needs to be fixed
        return $this->convertEncoding(substr($data, 1, $isCompressed ? 2 * $length : $length), $this->codePage ?: 'UTF-16LE');
    }

    /**
     * Convert UTF-16 string in compressed notation to uncompressed form. Only used for BIFF8.
     *
     * @param string $string
     *
     * @return string
     */
    private function uncompressByteString($string)
    {
        $uncompressedString = '';
        $strLen = strlen($string);
        for ($i = 0; $i < $strLen; ++$i) {
            $uncompressedString .= $string[$i] . "\0";
        }

        return $uncompressedString;
    }

    /**
     * Destructor
     */
    public function __destruct()
    {
        $this->spreadsheet = null;
        $this->workbook = null;
    }
}
