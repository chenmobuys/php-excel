<?php

namespace Excel\Reader;

use SplFileObject;
use Excel\Reader\Csv\Delimiter;
use Excel\Reader\Csv\RowIterator;
use Excel\Shared\Coordinate;

class Csv extends BaseReader
{
    const DEFAULT_FALLBACK_ENCODING = 'CP1252';
    const GUESS_ENCODING = 'guess';

    const UTF8_BOM = "\xef\xbb\xbf";
    const UTF16BE_BOM = "\xfe\xff";
    const UTF16BE_LF = "\x00\x0a";
    const UTF16LE_BOM = "\xff\xfe";
    const UTF16LE_LF = "\x0a\x00";
    const UTF32BE_BOM = "\x00\x00\xfe\xff";
    const UTF32BE_LF = "\x00\x00\x00\x0a";
    const UTF32LE_BOM = "\xff\xfe\x00\x00";
    const UTF32LE_LF = "\x0a\x00\x00\x00";

    /**
     * @var string|null
     */
    private $delimiter;

    /**
     * @var string
     */
    private $enclosure = '"';

    /**
     * @var string
     */
    private $escape = '\\';

    /**
     * @var \SplFileObject
     */
    private $splFileObject;

    /**
     * Input encoding.
     *
     * @var string
     */
    private $inputEncoding = self::GUESS_ENCODING;

    /**
     * Fallback encoding if guess strikes out.
     *
     * @var string
     */
    private $fallbackEncoding = self::DEFAULT_FALLBACK_ENCODING;

    /**
     * Load worksheet.
     * 
     * @param string $filename
     * @return void
     */
    public function load(string $filename): void
    {
        $this->splFileObject = $this->getSplFileObject($filename);
        $this->worksheetNames = $this->getWorksheetNamesByFilename($filename);
        $this->worksheetInfo = $this->getWorksheetInfoBySplFileObject($this->splFileObject);
    }

    /**
     * Determine whether the file is readable.
     * 
     * @param string $filename 
     * @return bool
     */
    public function isReadable(string $filename): bool
    {
        // Attempt to guess mimetype
        $type = mime_content_type($filename);

        $supportedTypes = [
            'application/csv',
            'text/csv',
            'text/plain',
            'inode/x-empty',
        ];

        return in_array($type, $supportedTypes, true);
    }

    /**
     * List worksheet names.
     * 
     * @param string $filename
     * @return array
     */
    public function listWorksheetNames(string $filename): array
    {
        return $this->getWorksheetNamesByFilename($filename);
    }

    /**
     * List worksheet info.
     * 
     * @param string $filename
     * @return array
     */
    public function listWorksheetInfo(string $filename): array
    {
        $splFileObject = $this->getSplFileObject($filename);
        return $this->getWorksheetInfoBySplFileObject($splFileObject);
    }

    /**
     * Get worksheet names by filename.
     * 
     * @param string $filename
     * 
     * @return array
     */
    private function getWorksheetNamesByFilename(string $filename): array
    {
        return [pathinfo($filename, PATHINFO_FILENAME)];
    }

    /**
     * Get worksheet info by SplFileObject.
     * 
     * @param \SplFileObject
     * 
     * @return array
     */
    private function getWorksheetInfoBySplFileObject(SplFileObject $splFileObject): array
    {
        $worksheetInfo = [];
        $worksheetInfo[0]['worksheetName'] = pathinfo($splFileObject->getFilename(), PATHINFO_FILENAME);
        $worksheetInfo[0]['lastColumnLetter'] = 'A';
        $worksheetInfo[0]['lastColumnIndex'] = 0;
        $worksheetInfo[0]['totalRows'] = 0;
        $worksheetInfo[0]['totalColumns'] = 0;

        while (!$splFileObject->eof()) {
            if (is_array($row = $splFileObject->current())) {
                ++$worksheetInfo[0]['totalRows'];
                $worksheetInfo[0]['lastColumnIndex'] = max($worksheetInfo[0]['lastColumnIndex'], count($row) - 1);
            }
            $splFileObject->next();
        }

        $worksheetInfo[0]['lastColumnLetter'] = Coordinate::columnLetterFromColumnIndex($worksheetInfo[0]['lastColumnIndex']);
        $worksheetInfo[0]['totalColumns'] = $worksheetInfo[0]['lastColumnIndex'] + 1;

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
    protected function getRowIterator(array $worksheetInfo,int $worksheetIndex, int $startRow = 1, int $endRow = null): BaseRowIterator
    {
        $this->splFileObject->rewind();
        $style = $this->getStyle();

        return new RowIterator($this->splFileObject, $style, 0, $startRow, $endRow);
    }

    /**
     * Set input encoding.
     * 
     * @param string $inputEncoding
     * @return self
     */
    public function setInputEncoding(string $inputEncoding): self
    {
        $this->inputEncoding = $inputEncoding;
        return $this;
    }

    /**
     * Set delimiter.
     * 
     * @param string $delimiter
     * @return self
     */
    public function setDelimiter(string $delimiter): self
    {
        $this->delimiter = $delimiter;
        return $this;
    }

    /**
     * Set enclosure.
     * 
     * @param string $enclosure
     * @return self
     */
    public function setEnclosure(string $enclosure): self
    {
        $this->enclosure = $enclosure;
        return $this;
    }

    /**
     * Set escape.
     * 
     * @param string $escape
     * @return self
     */
    public function setEscape(string $escape): self
    {
        $this->escape = $escape;
        return $this;
    }

    /**
     * Get input encoding.
     * 
     * @return string
     */
    public function getInputEncoding(): string
    {
        return $this->inputEncoding;
    }

    /**
     * Get delimiter.
     * 
     * @return string
     */
    public function getDelimiter(): string
    {
        return $this->delimiter;
    }

    /**
     * Get enclosure.
     * 
     * @return string
     */
    public function getEnclosure(): string
    {
        return $this->enclosure;
    }

    /**
     * Get escape.
     * 
     * @return string
     */
    public function getEscape(): string
    {
        return $this->escape;
    }

    /**
     * Get SplFileObject.
     * 
     * @param string $filename
     * @return \SplFileObject
     */
    private function getSplFileObject(string $filename): SplFileObject
    {
        $splFileObject = new SplFileObject($filename, 'r+b');

        if ($this->inputEncoding === self::GUESS_ENCODING) {
            $this->inputEncoding = self::guessEncoding($filename, $this->fallbackEncoding);
        }

        if ($this->inputEncoding !== 'UTF-8') {
            $contents = file_get_contents($filename);
            if ($contents !== false) {
                $contents = mb_convert_encoding($contents, 'UTF-8', $this->inputEncoding);
                $splFileObject->fwrite($contents);
                $splFileObject->rewind();
            }
        }

        if (!$this->delimiter) {
            $this->inferSeparator($splFileObject);
            $splFileObject->rewind();
        }

        $splFileObject->setFlags(SplFileObject::READ_CSV | SplFileObject::READ_AHEAD);
        $splFileObject->setCsvControl($this->delimiter, $this->enclosure, $this->escape);

        return $splFileObject;
    }

    /**
     * Infer spearator.
     * 
     * @param \SplFileObject $splFileObject
     */
    private function inferSeparator(SplFileObject $splFileObject): void
    {
        if ($line = $splFileObject->fgets() === false) {
            return;
        }

        if ((strlen(trim($line, "\r\n")) == 5) && (stripos($line, 'sep=') === 0)) {
            $this->delimiter = substr($line, 4, 1);
            return;
        }

        $splFileObject->rewind();

        $inferenceEngine = new Delimiter($splFileObject, $this->escape, $this->enclosure);

        // If number of lines is 0, nothing to infer : fall back to the default
        if ($inferenceEngine->linesCounted() === 0) {
            $this->delimiter = $inferenceEngine->getDefaultDelimiter();
            return;
        }

        $this->delimiter = $inferenceEngine->infer();

        // If no delimiter could be detected, fall back to the default
        if ($this->delimiter === null) {
            $this->delimiter = $inferenceEngine->getDefaultDelimiter();
            return;
        }
    }

    /**
     * Guess encoding test no bom.
     * 
     * @param string $encoding
     * @param string $contents
     * @param string $compare
     * @param string $setEncoding
     * @return void
     */
    private static function guessEncodingTestNoBom(string &$encoding, string &$contents, string $compare, string $setEncoding): void
    {
        if ($encoding === '') {
            $pos = strpos($contents, $compare);
            if ($pos !== false && $pos % strlen($compare) === 0) {
                $encoding = $setEncoding;
            }
        }
    }

    /**
     * Guess encoding no bom.
     * 
     * @param string $filename
     * @return void
     */
    private static function guessEncodingNoBom(string $filename): string
    {
        $encoding = '';
        $contents = file_get_contents($filename);
        self::guessEncodingTestNoBom($encoding, $contents, self::UTF32BE_LF, 'UTF-32BE');
        self::guessEncodingTestNoBom($encoding, $contents, self::UTF32LE_LF, 'UTF-32LE');
        self::guessEncodingTestNoBom($encoding, $contents, self::UTF16BE_LF, 'UTF-16BE');
        self::guessEncodingTestNoBom($encoding, $contents, self::UTF16LE_LF, 'UTF-16LE');
        if ($encoding === '' && preg_match('//u', $contents) === 1) {
            $encoding = 'UTF-8';
        }

        return $encoding;
    }

    /**
     * Guess encoding test with bom.
     * 
     * @param string $encoding
     * @param string $first4
     * @param string $compare
     * @param string $setEncoding
     * @return void
     */
    private static function guessEncodingTestBom(string &$encoding, string $first4, string $compare, string $setEncoding): void
    {
        if ($encoding === '') {
            if ($compare === substr($first4, 0, strlen($compare))) {
                $encoding = $setEncoding;
            }
        }
    }

    /**
     * Guess encoding with bom.
     * 
     * @param string $filename
     * @return void
     */
    private static function guessEncodingBom(string $filename): string
    {
        $encoding = '';
        $first4 = file_get_contents($filename, false, null, 0, 4);
        if ($first4 !== false) {
            self::guessEncodingTestBom($encoding, $first4, self::UTF8_BOM, 'UTF-8');
            self::guessEncodingTestBom($encoding, $first4, self::UTF16BE_BOM, 'UTF-16BE');
            self::guessEncodingTestBom($encoding, $first4, self::UTF32BE_BOM, 'UTF-32BE');
            self::guessEncodingTestBom($encoding, $first4, self::UTF32LE_BOM, 'UTF-32LE');
            self::guessEncodingTestBom($encoding, $first4, self::UTF16LE_BOM, 'UTF-16LE');
        }

        return $encoding;
    }

    /**
     * Guess encoding by mb.
     * 
     * @param string $filename
     * @return string
     */
    private static function guessEncodingMb($filename): string
    {
        return mb_detect_encoding(file_get_contents($filename), mb_list_encodings());
    }

    /**
     * Guess encoding.
     * 
     * @param string $filename
     * @param string $dflt
     * @return string
     */
    public static function guessEncoding(string $filename, string $dflt = self::DEFAULT_FALLBACK_ENCODING): string
    {
        $encoding = self::guessEncodingBom($filename);
        if ($encoding === '') {
            $encoding = self::guessEncodingNoBom($filename);
        }
        if ($encoding === '') {
            $encoding = self::guessEncodingMb($filename);
        }

        return ($encoding === '') ? $dflt : $encoding;
    }

    public function __destruct()
    {
        $this->splFileObject = null;
    }
}
