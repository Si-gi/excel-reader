<?php

declare(strict_types=1);

namespace Sigi\ExcelReader\Objects;

class CellCollection implements \IteratorAggregate, \Countable, \ArrayAccess
{
    private array $cells = [];

    public function add(Cell $cell): void
    {
        $this->cells[] = $cell;
    }

    public function get(int $index): ?Cell
    {
        return $this->cells[$index] ?? null;
    }

    public function count(): int
    {
        return count($this->cells);
    }

    public function getIterator(): \ArrayIterator
    {
        return new \ArrayIterator($this->cells);
    }

    public function toArray(): array
    {
        return array_map(fn(Cell $cell) => $cell->getValue(), $this->cells);
    }

    public function first(): ?Cell
    {
        return $this->cells[0] ?? null;
    }

    public function last(): ?Cell
    {
        return $this->cells[count($this->cells) - 1] ?? null;
    }

    // ArrayAccess
    public function offsetExists($offset): bool
    {
        return isset($this->cells[$offset]);
    }

    public function offsetGet($offset): ?Cell
    {
        return $this->cells[$offset] ?? null;
    }

    public function offsetSet($offset, $value): void
    {
        if ($offset === null) {
            $this->cells[] = $value;
        } else {
            $this->cells[$offset] = $value;
        }
    }

    public function offsetUnset($offset): void
    {
        unset($this->cells[$offset]);
    }
}