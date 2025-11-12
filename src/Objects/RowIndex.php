<?php

declare(strict_types=1);

namespace Sigi\ExcelReader\Objects;

/**
 * Store index of row and their start and end for pointer
 */
class RowIndex {
     private array $index = [];

    public function add(int $rowNumber, int $startOffset, int $length): void
    {
        $this->index[$rowNumber] = [
            'start' => $startOffset,
            'length' => $length,
        ];
    }

    public function get(int $rowNumber): ?array
    {
        return $this->index[$rowNumber] ?? null;
    }

    public function count(): int
    {
        return count($this->index);
    }

    public function getMemoryUsage(): int
    {
        return count($this->index) * 40;
    }
}