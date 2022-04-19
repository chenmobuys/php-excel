<?php

namespace Excel;

use Excel\Reader\BaseReader;
use Excel\Exception\SpreadsheetException;

/**
 * Main class for spreadsheet reading
 *
 * @author Chenmobuys
 */
class SpreadsheetFactory
{
    const EXCEL_CSV = 'Csv';
    const EXCEL_ODS = 'Ods';
    const EXCEL_XLS = 'Xls';
    const EXCEL_XLSX = 'Xlsx';
    const EXCEL_XML = 'Xml';
    const EXCEL_SLK = 'Slk';
    const EXCEL_GNUMERIC = 'Gnumeric';
    const EXCEL_HTML = 'Html';

    private static $readers = [
        self::EXCEL_CSV => Reader\Csv::class,
        self::EXCEL_ODS => Reader\Ods::class,
        self::EXCEL_XLS => Reader\Xls::class,
        self::EXCEL_XLSX => Reader\Xlsx::class,
        // TODO next to
        // 'Xml' => Readers\Xml::class,
        // 'Slk' => Reader\Slk::class,
        // 'Gnumeric' => Reader\Gnumeric::class,
        // 'Html' => Reader\Html::class,
    ];

    /**
     * Create Reader For ReaderType
     *
     * @param string $readerType
     *
     * @return BaseReader
     */
    public static function createReader($readerType)
    {
        if (!isset(self::$readers[$readerType])) {
            throw new SpreadsheetException("No reader found for type $readerType");
        }

        // Instantiate reader
        $className = self::$readers[$readerType];

        return new $className();
    }

    /**
     * Create Reader For File
     *
     * @param string $filename
     * @param bool $load
     *
     * @return BaseReader
     */
    public static function createReaderForFile($filename, $load = false)
    {
        if (!is_file($filename)) {
            throw new SpreadsheetException('File "' . $filename . '" does not exist.');
        }

        if (!is_readable($filename)) {
            throw new SpreadsheetException('Could not open "' . $filename . '" for reading.');
        }

        $guessedReader = self::getReaderTypeFromExtension($filename);

        if ($guessedReader !== null) {
            $reader = self::createReader($guessedReader);

            // Let's see if we are lucky
            if ($reader->isReadable($filename)) {
                $load && $reader->load($filename);
                return $reader;
            }
        }

        foreach (self::$readers as $type => $class) {
            //    Ignore our original guess, we know that won't work
            if ($type !== $guessedReader) {
                $reader = self::createReader($type);
                if ($reader->isReadable($filename)) {
                    $load && $reader->load($filename);
                    return $reader;
                }
            }
        }

        throw new SpreadsheetException('Unable to identify a reader for this file');
    }

    /**
     * Get ReaderType from Extension
     *
     * @param $filename
     *
     * @return string $readerType
     *
     */
    private static function getReaderTypeFromExtension($filename)
    {
        $extension = strtolower(pathinfo($filename, PATHINFO_EXTENSION));

        if (is_null($extension)) {
            return null;
        }

        switch (strtolower($extension)) {
            case 'xlsx': // Excel (OfficeOpenXML) Spreadsheet
            case 'xlsm': // Excel (OfficeOpenXML) Macro Spreadsheet (macros will be discarded)
            case 'xltx': // Excel (OfficeOpenXML) Template
            case 'xltm': // Excel (OfficeOpenXML) Macro Template (macros will be discarded)
                return self::EXCEL_XLSX;
            case 'xls': // Excel (BIFF) Spreadsheet
            case 'xlt': // Excel (BIFF) Template
                return self::EXCEL_XLS;
            case 'ods': // Open/Libre Offic Calc
            case 'ots': // Open/Libre Offic Calc Template
                return self::EXCEL_ODS;
            case 'slk':
                return self::EXCEL_SLK;
            case 'xml': // Excel 2003 SpreadSheetML
                return self::EXCEL_XML;
            case 'gnumeric':
                return self::EXCEL_GNUMERIC;
            case 'htm':
            case 'html':
                return self::EXCEL_HTML;
            case 'csv':
            case 'tsv':
                return self::EXCEL_CSV;
            default:
                return null;
        }
    }

    /**
     * Register a reader with it's type and class name.
     *
     * @param string $readerType
     * @param string $readerClass
     */
    public static function registerReader($readerType, $readerClass)
    {
        if (!is_a($readerClass, BaseReader::class, true)) {
            throw new SpreadsheetException('Registered readers must implement ' . BaseReader::class);
        }

        self::$readers[$readerType] = $readerClass;
    }
}
