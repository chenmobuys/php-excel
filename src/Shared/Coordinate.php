<?php

namespace Excel\Shared;

class Coordinate
{
    public static function columnIndexFromColumnLetter($letter)
    {
        $letter = strtoupper($letter);

        $result = 0;
        for ($i = strlen($letter) - 1, $j = 0; $i >= 0; $i--, $j++) {
            $ord = ord($letter[$i]) - 64;
            if ($ord > 26) {
                return false;
            }
            $result += $ord * pow(26, $j);
        }
        return $result - 1;
    }

    public static function columnLetterFromColumnIndex($index)
    {
        $index++;
        $letter = null;
        do {
            $characterValue = ($index % 26) ?: 26;
            $index = ($index - $characterValue) / 26;
            $letter = chr($characterValue + 64) . ($letter ?: '');
        } while ($index > 0);
        return $letter;
    }
}
