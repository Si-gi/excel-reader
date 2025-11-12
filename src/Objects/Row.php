<?php

declare(strict_types=1);

namespace Sigi\ExcelReader\Objects;

use Sigi\ExcelReader\Reader\SharedStringsReader;


class Row
{
    private ?CellCollection $cells = null;
    private bool $parsed = false;
    private string $rowXml;

    public function __construct(
        private Sheet $sheet,
        private int $number,
        string $rowXml
    ) {
        $this->rowXml = $rowXml;
    }
    public function getNumber(): int
    {
        return $this->number;
    }

    public function getCells(): CellCollection
    {
        if (!$this->parsed) {
            $this->parse();
        }

        return $this->cells;
    }
    public function getCell(int $index): ?Cell
    {
        return $this->getCells()->get($index);
    }

    public function contains(mixed $value, bool $strict = true): bool
    {
        foreach ($this->getCells() as $cell) {
            if ($strict) {
                if ($cell->getValue() === $value) {
                    return true;
                }
            } else {
                if ($cell->getValue() == $value) {
                    return true;
                }
            }
        }

        return false;
    }
    // Check if a value of a cell match a regex
    public function matches(string $pattern): bool
    {
        foreach ($this->getCells() as $cell) {
            if (preg_match($pattern, (string)$cell->getValue())) {
                return true;
            }
        }

        return false;
    }

    /**
    * Clone la ligne pour la retourner (détache de Sheet)
    */
    public function detach(): self
    {
        $clone = new self(
            sheet: $this->sheet,
            number: $this->number,
            rowXml: ''

        );
        
        if ($this->parsed) {
            $clone->cells = $this->cells;
            $clone->parsed = true;
        }
        
        return $clone;
    }

    /**
    * Libère la mémoire des cellules
    */
    public function free(): void
    {
        $this->cells = null;
        $this->parsed = false;
    }

    public function __destruct()
    {
        
    }
    public function toArray(): array
    {
        return $this->getCells()->toArray();
    }

    /**
     * Parse XML of The line
     */
    private function parse(): void
    {
        $this->cells = new CellCollection();
        $this->parseCells($this->rowXml);
        $this->parsed = true;

        // free XML after parsing
        $this->rowXml = '';
    }


    private function parseCells(string $rowXml): void
    {
        $sharedStringsReader = $this->sheet->getReader()->getSharedStringsReader();

        preg_match_all('/<c[^>]*>.*?<\/c>/s', $rowXml, $cellMatches);

        $columnIndex = 0;
        foreach ($cellMatches[0] as $cellXml) {


            $cell = $this->parseCell($cellXml, $sharedStringsReader, $columnIndex);


            $this->cells->add($cell);
            $columnIndex++;
        }
    }

    private function parseCell(string $cellXml, SharedStringsReader  $sharedStringsReader, int $columnIndex): Cell
    {
        // Extraire la référence de la cellule (ex: A1, B2)
        preg_match('/r="([^"]+)"/', $cellXml, $refMatch);
        $reference = $refMatch[1] ?? '';

        // Extraire le type
        $isSharedString = str_contains($cellXml, 't="s"');
        $isInlineString = str_contains($cellXml, 't="inlineStr"');

        $value = '';

        if ($isInlineString) {
            if (preg_match('/<t>([^<]*)<\/t>/', $cellXml, $match)) {
                $value = $match[1];
            }
        } elseif (preg_match('/<v>([^<]+)<\/v>/', $cellXml, $match)) {
            $value = $match[1];

            if ($isSharedString) {
                $index = (int)$value;
                $value = $sharedStringsReader->get($index);
            }
        }

        return new Cell(
            value: $value,
            columnIndex: $columnIndex,
            reference: $reference
        );
    }
}