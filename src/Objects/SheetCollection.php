<?php

declare(strict_types=1);

namespace Sigi\ExcelReader\Objects;

use ArrayAccess;
use Countable;
use IteratorAggregate;
use ArrayIterator;

class SheetCollection implements ArrayAccess, Countable, IteratorAggregate
{
    private array $sheets = [];

    public function __construct(array $sheets = [])
    {
        $this->sheets = $sheets;
    }

    public function add(Sheet $sheet): void
    {
        $this->sheets[] = $sheet;
    }

    public function set(string $key, Sheet $sheet): void
    {
        $this->sheets[$key] = $sheet;
    }

    public function get(string $key): ?Sheet
    {
        return $this->sheets[$key] ?? null;
    }

    public function remove(string $key): void
    {
        unset($this->sheets[$key]);
    }

    public function contains(Sheet $sheet): bool
    {
        return in_array($sheet, $this->sheets, true);
    }

    public function isEmpty(): bool
    {
        return empty($this->sheets);
    }

    public function clear(): void
    {
        $this->sheets = [];
    }

    public function toArray(): array
    {
        return $this->sheets;
    }

    public function first(): ?Sheet
    {
        return reset($this->sheets) ?: null;
    }

    public function last(): ?Sheet
    {
        return end($this->sheets) ?: null;
    }

    public function filter(callable $callback): self
    {
        return new self(array_filter($this->sheets, $callback));
    }

    public function map(callable $callback): array
    {
        return array_map($callback, $this->sheets);
    }

    public function findByName(string $name): ?Sheet
    {
        foreach ($this->sheets as $sheet) {
            if ($sheet->getName() === $name) {
                return $sheet;
            }
        }
        return null;
    }

    // Interface ArrayAccess
    public function offsetExists($offset): bool
    {
        return isset($this->sheets[$offset]);
    }

    public function offsetGet($offset): ?Sheet
    {
        return $this->sheets[$offset] ?? null;
    }

    public function offsetSet($offset, $value): void
    {
        if ($offset === null) {
            $this->sheets[] = $value;
        } else {
            $this->sheets[$offset] = $value;
        }
    }

    public function offsetUnset($offset): void
    {
        unset($this->sheets[$offset]);
    }

    // Interface Countable
    public function count(): int
    {
        return count($this->sheets);
    }

    // Interface IteratorAggregate
    public function getIterator(): ArrayIterator
    {
        return new ArrayIterator($this->sheets);
    }
}