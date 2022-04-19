<?php

namespace ExcelTests\Feature;

use Excel\Reader\BaseReader;
use Excel\Reader\BaseRowIterator;

class FooReader extends BaseReader
{
    public function load(string $filename): void
    {
    }

    public function isReadable(string $filename): bool
    {
        return true;
    }

    public function listWorksheetNames(string $filename): array
    {
        return [];
    }

    public function listWorksheetInfo(string $filename): array
    {
        return [];
    }

    protected function getRowIterator(array $worksheetInfo, int $worksheetIndex, int $startRow = 1, int $endRow = null): BaseRowIterator
    {
        return new BaseRowIterator($this->style, $worksheetIndex, $startRow, $endRow);
    }
}
