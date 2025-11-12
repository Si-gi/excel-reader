<?php

declare(strict_types=1);

namespace Sigi\ExcelReader\Objects;

class RowCollection implements \IteratorAggregate, \Countable
{
    private array $rows = [];

    public function add(Row $row): void
    {
        $this->rows[] = $row;
    }

    public function get(int $index): ?Row
    {
        return $this->rows[$index] ?? null;
    }

    public function count(): int
    {
        return count($this->rows);
    }

    public function getIterator(): \ArrayIterator
    {
        return new \ArrayIterator($this->rows);
    }

    public function toArray(): array
    {
        return $this->rows;
    }

    public function first(): ?Row
    {
        return $this->rows[0] ?? null;
    }

    public function last(): ?Row
    {
        return $this->rows[count($this->rows) - 1] ?? null;
    }
}