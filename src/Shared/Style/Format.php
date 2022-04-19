<?php

namespace Excel\Shared\Style;

use DateTime;
use DateInterval;
use Excel\Shared\Cell;

class Format extends SharedComponent
{
    // Pre-defined formats
    const FORMAT_GENERAL = 'General';
    const FORMAT_TEXT = '@';
    const FORMAT_POSTCODE = '000000';

    const CALENDAR_WINDOWS_1900 = 1900; //    Base date of 1st Jan 1900 = 1.0
    const CALENDAR_MAC_1904 = 1904;     //    Base date of 2nd Jan 1904 = 1.0

    /**
     * Parsed format code cache
     * 
     * @var array
     */
    private static $parsedCache = [];

    /**
     * Calendar
     * 
     * @var int
     */
    private static $calendar = self::CALENDAR_WINDOWS_1900;

    /**
     * Set Calendar.
     * 
     * @var int $calendar
     */
    public function setCalendar(int $calendar): void
    {
        self::$calendar = $calendar;
    }

    /**
     * Get Calendar.
     * 
     * @return int
     */
    public function getCalendar(): int
    {
        return self::$calendar;
    }

    /**
     * Get value.
     * 
     * @param \Excel\Shared\Cell $cell
     * @return string
     */
    public function getValue(Cell $cell): string
    {
        $value = $cell->getValue();

        $formatCode = $this->getFormatCode($cell);

        // General Format
        if ($formatCode === self::FORMAT_GENERAL) {
            $value = (string) $value;
        }
        // Text Format
        elseif ($formatCode === self::FORMAT_TEXT) {
            $value = (string) $value;
        }
        // Postcode Format
        elseif ($formatCode === self::FORMAT_POSTCODE) {
            $value = (string) $value;
        }
        // Number Format
        elseif (preg_match('/^((#,##)?0)\.?(0*)$/', $formatCode)) {
            $precision = self::getFormatPrecision($formatCode);
            $thousandsSeparator = strpos($formatCode, ',') !== false ? ',' : '';
            $value = number_format($value, $precision, '.', $thousandsSeparator);
        }
        // Number F-E Format
        elseif (preg_match('/0\.0*E0+/', $formatCode)) {
            $value = strtoupper($value);
            if (strpos($value, 'E') !== false) {
                list($first, $second) = explode('E', strtoupper($value));
                $precision = self::getFormatPrecision($first);
                $precisionSecond = abs($second);
                $value = bcmul($first, bcpow(10, $second, $precisionSecond), $precision + $precisionSecond);
            }
        }
        // Currency Format
        elseif (preg_match('/^((?<currency>(￥|€|\$|US\$)?)#,##0)\.?(0*)$/', $formatCode, $matches)) {
            $precision = self::getFormatPrecision($formatCode);
            $currency = $matches['currency'];
            $thousandsSeparator = strpos($formatCode, ',') !== false ? ',' : '';
            $value = $currency . number_format($value, $precision, '.', $thousandsSeparator);
        }
        // Percent Format
        elseif (preg_match('/^0(\.)?0*%$/', $formatCode)) {
            $precision = self::getFormatPrecision(trim($formatCode, '%'));
            $value = bcmul($value, 100, $precision) . '%';
        }
        // Fraction Format
        elseif (preg_match('/^#\s(\?)*\/(\?)*$/', $formatCode)) {
            // Do nothing.
        }
        // Datetime Format
        elseif (preg_match('/^(\[\$[[:alpha:]]*-[0-9A-F]*\])*[hmsdy]/i', $formatCode)) {
            if (self::$calendar === self::CALENDAR_WINDOWS_1900) {
                $baseDate = ($value < 1) ? new DateTime('1970-01-01')
                    : (($value < 60) ? new DateTime('1899-12-31') : new DateTime('1899-12-30'));
            } else {
                $baseDate = new DateTime('1904-01-01');
            }

            $dateFormat = strtolower($formatCode);
            $dateFormat = strtr($dateFormat, self::$dateReplacements);
            if (strpos($dateFormat, 'A') === false) {
                $dateFormat = strtr($dateFormat, self::$dateReplacements24);
            } else {
                $dateFormat = strtr($dateFormat, self::$dateReplacements12);
            }

            $days = (int) $value;
            $seconds = (int) (($value - $days) * 86400);
            $dateInterval = new DateInterval(sprintf('P%sD%s', $days, $seconds ? 'T' . $seconds . 'S' : ''));
            $value  = $baseDate->add($dateInterval)->format($dateFormat);
        }
        // TODO Other Format

        return $value;
    }

    /**
     * Get format code.
     * 
     * @param \Excel\Shared\Cell $cell
     * @return string
     */
    private function getFormatCode(Cell $cell): string
    {
        $formatCode = self::FORMAT_GENERAL;

        if (!$cell->getXfIndex()) {
            return $formatCode;
        }

        $formatIndex = $this->style->getXF()->getValue($cell);

        if (!$formatIndex) {
            return $formatCode;
        }

        if (self::$parsedCache[$formatIndex] ?? false) {
            return self::$parsedCache[$formatIndex];
        }

        if (self::$builtinFormats[$formatIndex] ?? false) {
            $formatCode = self::$builtinFormats[$formatIndex];
        }

        if ($this->items[$formatIndex] ?? false) {
            $formatCode = $this->items[$formatIndex];
        }

        $formatCodes = explode(';', $formatCode);
        $formatCode = current($formatCodes);
        switch (count($formatCodes)) {
            case 2:
                $cell->getValue() < 0 && $formatCode = $formatCodes[1];
                break;
            case 3:
            case 4:
                $cell->getValue() < 0 && $formatCode = $formatCodes[1];
                $cell->getValue() == 0 && $formatCode = $formatCodes[2];
                break;
        }

        // Stripping colors
        $formatCode = trim(preg_replace('/^\[[[:alpha:]]+\]/i', '', $formatCode));
        // Removing skipped characters
        $formatCode = preg_replace('/_/', '', $formatCode);
        // Removing unnecessary escaping
        $formatCode = preg_replace("/\\\\/", '', $formatCode);
        // Removing string quotes
        $formatCode = str_replace(['"', '*'], '', $formatCode);
        // Removing plus or minus
        $formatCode = str_replace(['+', '-'], '', $formatCode);

        return $formatCode;
    }

    /**
     * Get format precision.
     * 
     * @param string
     */
    private static function getFormatPrecision(string $value): int
    {
        return strlen($value) - (strpos($value, '.') ?: 0) - 1;
    }

    /**
     * @var array $builtinFormats
     */
    private static $builtinFormats = [
        0 => '',
        1 => '0',
        2 => '0.00',
        3 => '#,##0',
        4 => '#,##0.00',

        9 => '0%',
        10 => '0.00%',
        11 => '0.00E+00',
        12 => '# ?/?',
        13 => '# ??/??',
        14 => 'mm-dd-yy',
        15 => 'd-mmm-yy',
        16 => 'd-mmm',
        17 => 'mmm-yy',
        18 => 'h:mm AM/PM',
        19 => 'h:mm:ss AM/PM',
        20 => 'h:mm',
        21 => 'h:mm:ss',
        22 => 'm/d/yy h:mm',

        37 => '#,##0 ;(#,##0)',
        38 => '#,##0 ;[Red](#,##0)',
        39 => '#,##0.00;(#,##0.00)',
        40 => '#,##0.00;[Red](#,##0.00)',

        44 => '_("$"* #,##0.00_);_("$"* \(#,##0.00\);_("$"* "-"??_);_(@_)',
        45 => 'mm:ss',
        46 => '[h]:mm:ss',
        47 => 'mmss.0',
        48 => '##0.0E+0',
        49 => '@',

        // CHT & CHS
        27 => '[$-404]e/m/d',
        30 => 'm/d/yy',
        36 => '[$-404]e/m/d',
        50 => '[$-404]e/m/d',
        57 => '[$-404]e/m/d',

        // THA
        59 => 't0',
        60 => 't0.00',
        61 => 't#,##0',
        62 => 't#,##0.00',
        67 => 't0%',
        68 => 't0.00%',
        69 => 't# ?/?',
        70 => 't# ??/??',

        // JPN
        28 => '[$-411]ggge"年"m"月"d"日"',
        29 => '[$-411]ggge"年"m"月"d"日"',
        31 => 'yyyy"年"m"月"d"日"',
        32 => 'h"時"mm"分"',
        33 => 'h"時"mm"分"ss"秒"',
        34 => 'yyyy"年"m"月"',
        35 => 'm"月"d"日"',
        51 => '[$-411]ggge"年"m"月"d"日"',
        52 => 'yyyy"年"m"月"',
        53 => 'm"月"d"日"',
        54 => '[$-411]ggge"年"m"月"d"日"',
        55 => 'yyyy"年"m"月"',
        56 => 'm"月"d"日"',
        58 => '[$-411]ggge"年"m"月"d"日"',
    ];

    /**
     * @var array $dateReplacements
     */
    private static $dateReplacements = [
        '\\' => '',
        'am/pm' => 'A',
        'yyyy' => 'Y',
        'yy' => 'y',
        'mmmmm' => 'M',
        'mmmm' => 'F',
        'mmm' => 'M',
        ':mm' => ':i',
        'mm' => 'm',
        'm' => 'n',
        'dddd' => 'l',
        'ddd' => 'D',
        'dd' => 'd',
        'd' => 'j',
        'ss' => 's',
        '.s' => '',
        '12H' => []
    ];

    /**
     * @var array $dateReplacements12
     */
    private static $dateReplacements12 = [
        'hh' => 'h',
        'h' => 'G',
    ];

    /**
     * @var array $dateReplacements24
     */
    private static $dateReplacements24 = [
        'hh' => 'H',
        'h' => 'G'
    ];
}
