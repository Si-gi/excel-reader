<?php

namespace Sigi\ExcelReader\Objects;

use XMLReader;
use ZipArchive;
use Sigi\ExcelReader\Reader\Reader;
use Sigi\ExcelReader\Detector\ExcelStructure;
use Sigi\ExcelReader\Reader\SharedStringsReader;


class Sheet
{
    public function __construct(
        private ZipArchive $zip,
        private ExcelStructure $structure,
        private ?SharedStringsReader $sharedStringsReader,
        private string $name,
        private int $index,
        private string $xmlPath
    ) {
    }

    public function getName(): string
    {
        return $this->name;
    }

    public function getIndex(): int
    {
        return $this->index;
    }

    public function getXmlPath(): string
    {
        return $this->xmlPath;
    }

    public function getZip(): ZipArchive
    {
        return $this->zip;
    }

    public function getStructure(): ExcelStructure
    {
        return $this->structure;
    }

    public function getSharedStringsReader(): ?SharedStringsReader
    {
        return $this->sharedStringsReader;
    }

    /**
     * Itère sur les lignes
     */
    public function getRowIterator(): \Generator
    {
        var_dump($this->structure);
        
        $stream = $this->zip->getStream("xl/".$this->xmlPath);

        if ($stream === false) {
            throw new \RuntimeException("Could not open stream for {$this->xmlPath}");
        }

        $buffer = '';
        $rowNumber = 0;

        // Utiliser le pattern de row détecté par la structure
        $rowPattern = $this->getRowPattern();

        while (!feof($stream)) {
            $chunk = fread($stream, 8192);
            $buffer .= $chunk;

            while (preg_match($rowPattern, $buffer, $match, PREG_OFFSET_CAPTURE)) {
                $rowNumber++;
                $rowXml = $match[0][0];
                $matchStart = $match[0][1];

                yield new Row(
                    sheet: $this,
                    number: $rowNumber,
                    rowXml: $rowXml
                );

                $buffer = substr($buffer, $matchStart + strlen($rowXml));
            }

            if (strlen($buffer) > 50000) {
                $buffer = substr($buffer, -10240);
            }
        }

        fclose($stream);
    }

    /**
     * Retourne le pattern regex pour matcher les lignes
     * (peut être adapté selon la structure détectée)
     */
    private function getRowPattern(): string
    {
        // Pattern de base qui fonctionne avec la plupart des namespaces
        return '/<(?:[^:]+:)?row[^>]*>.*?<\/(?:[^:]+:)?row>/s';
    }

    public function findRow(mixed $value, bool $strict = true): ?Row
    {
        $count = 0;

        foreach ($this->getRowIterator() as $row) {
            $count++;

            if ($row->contains($value, $strict)) {
                return $row;
            }

            unset($row);

            if ($count % 5000 === 0) {
                gc_collect_cycles();
                $this->sharedStringsReader?->trimCache(100);
            }
        }

        return null;
    }

    public function findRowByPattern(string $pattern): ?Row
    {
        $count = 0;

        foreach ($this->getRowIterator() as $row) {
            $count++;

            if ($row->matches($pattern)) {
                return $row;
            }

            unset($row);

            if ($count % 5000 === 0) {
                gc_collect_cycles();
                $this->sharedStringsReader?->trimCache(100);
            }
        }

        return null;
    }

    public function filter(callable $callback, ?int $limit = null): RowCollection
    {
        $results = new RowCollection();
        $count = 0;

        foreach ($this->getRowIterator() as $row) {
            $count++;

            if ($callback($row)) {
                $results->add($row);

                if ($limit !== null && count($results) >= $limit) {
                    break;
                }
            } else {
                unset($row);
            }

            if ($count % 5000 === 0) {
                gc_collect_cycles();
                $this->sharedStringsReader?->trimCache(100);
            }
        }

        return $results;
    }
}