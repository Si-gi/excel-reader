<?php

declare(strict_types=1);

namespace Sigi\ExcelReader\Objects;

class Cell
{
    public function __construct(
        private mixed $value,
        private int $columnIndex,
        private string $reference = ''
    ) {
    }

    public function getValue(): mixed
    {
        return $this->value;
    }

    public function getColumnIndex(): int
    {
        return $this->columnIndex;
    }

    public function getReference(): string
    {
        return $this->reference;
    }

    public function __toString(): string
    {
        return (string)$this->value;
    }
}