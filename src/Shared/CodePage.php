<?php

namespace Excel\Shared;

class CodePage
{
    const DEFUALT_CODE_PAGE = 'CP1252';

    private static $codePageIdentifiers = [
        367 => 'ASCII',     // (ASCII)
        437 => 'CP437',     // (US)
        720 => 'CP720',     // (OEM Arabic)
        737 => 'CP737',     // (Greek)
        775 => 'CP775',     // (Baltic)
        850 => 'CP850',     // (Latin I)
        852 => 'CP852',     // (Latin II (Central European))
        855 => 'CP855',     // (Cyrillic)
        857 => 'CP857',     // (Turkish)
        858 => 'CP858',     // (Multilingual Latin I with Euro)
        860 => 'CP860',     // (Portuguese)
        861 => 'CP861',     // (Icelandic)
        862 => 'CP862',     // (Hebrew)
        863 => 'CP863',     // (Canadian (French))
        864 => 'CP864',     // (Arabic)
        865 => 'CP865',     // (Nordic)
        866 => 'CP866',     // (Cyrillic (Russian))
        869 => 'CP869',     // (Greek (Modern))
        874 => 'CP874',     // (Thai)
        932 => 'CP932',     // (Japanese Shift-JIS)
        936 => 'CP936',     // (Chinese Simplified GBK)
        949 => 'CP949',     // (Korean' (Wansung))
        950 => 'CP950',     // (Chinese Traditional BIG5)
        1200 => 'UTF-16LE', // (BIFF8)
        1250 => 'CP1250',   // (Latin II) (Central European)
        1251 => 'CP1251',   // (Cyrillic)
        1252 => 'CP1252',   // (Latin I) (BIFF4-BIFF5)
        1253 => 'CP1253',   // (Greek)
        1254 => 'CP1254',   // (Turkish)
        1255 => 'CP1255',   // (Hebrew)
        1256 => 'CP1256',   // (Arabic)
        1257 => 'CP1257',   // (Baltic)
        1258 => 'CP1258',   // (Vietnamese)
        1361 => 'CP1361',   // (Korean' (Johab))
        10000 => 'MAC',     // (Apple Roman)
        32768 => 'MAC',     // (Apple Roman)
        32769 => 'CP1252',  // (Latin I) (BIFF2-BIFF3)
        65000 => 'UTF-7',   // (Unicode (UTF-7))
        65001 => 'UTF-8',   // (Unicode (UTF-8))
    ];

    public static function numberToName(int $number)
    {
        return self::$codePageIdentifiers[$number] ?? self::DEFUALT_CODE_PAGE;
    }
}
