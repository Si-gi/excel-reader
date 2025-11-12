<?php

namespace Sigi\ExcelReader\Objects;

use XMLReader;
use ZipArchive;
use Sigi\ExcelReader\Reader\Reader;


class Sheet
{
    private ?RowIndex $rowIndex = null;
    private bool $indexed = false;
    private ?string $tempXmlFile = null;

    public function __construct(
        private Reader $reader,
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

    public function getReader(): Reader
    {
        return $this->reader;
    }

    /**
     * Search a line by cell value
     */
    public function findRow(mixed $value, bool $strict = true): ?Row
    {
        $count = 0;
        /** @var Row row */
        foreach ($this->getRowIterator() as $row) {
            $checkContain = microtime(true);
            $count++;

            if ($row->contains($value, $strict)) {
                return $row;
            }
            // // GC périodique
            // if ($count % 5000 === 0) {
            //     gc_collect_cycles();
            //     $this->reader->getSharedStringsReader()->trimCache(100);
            // }
        }

        return null;
    }

    /**
     * Search a line by cell value with regex
     */
    public function findRowByPattern(string $pattern): ?Row
    {
        $count = 0;
        $gcInterval = 1000;
        foreach ($this->getRowIterator() as $row) {
            if ($row->matches($pattern)) {
                $result = $row->detach();
                unset($row);
                return $result;
            }
            $row->free();
            unset($row);

            $count++;
            if ($count % $gcInterval === 0) {
                gc_collect_cycles();
                $this->reader->getSharedStringsReader()->clearCache();
            }
        }

        return null;
    }

    /**
     * Search every lines with callback
     */
    public function findRows(callable $callback, ?int $limit = null): RowCollection
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
                $this->reader->getSharedStringsReader()->trimCache(100);
            }
        }

        return $results;
    }

    /**
     * Search many lines by regex
     */
    public function findRowsByPattern(string $pattern, ?int $limit = null): RowCollection
    {
        return $this->findRows(fn(Row $row) => $row->matches($pattern), $limit);
    }

    /**
     * Alias of @findRows
     */
    public function filter(callable $callback): RowCollection
    {
        return $this->findRows($callback);
    }

    /**
     * Get an excel row by index
     */
    public function getRow(int $rowNumber): ?Row
    {
        if (!$this->indexed) {
            $this->buildIndex();
        }

        $position = $this->rowIndex->get($rowNumber);

        if ($position === null) {
            return null;
        }

        return new Row(
            sheet: $this,
            number: $rowNumber,
            rowXml: ''
        );
    }

    public function readRowContent(int $offset, int $length): string
    {
        $tempFile = $this->ensureTempFile();

        $handle = fopen($tempFile, 'r');
        if ($handle === false) {
            throw new \RuntimeException("Could not open temp file");
        }

        fseek($handle, $offset);
        $content = fread($handle, $length);
        fclose($handle);

        return $content;
    }

    /**
     * extract XML in stream (without loading in memory/RAM)
     */
    private function ensureTempFile(): string
    {
        if ($this->tempXmlFile !== null && file_exists($this->tempXmlFile)) {
            return $this->tempXmlFile;
        }

        $this->tempXmlFile = sys_get_temp_dir() . '/excel_sheet_' . uniqid() . '.xml';

        // Open file from ZIP
        $stream = $this->reader->getZip()->getStream($this->xmlPath);
        
        if ($stream === false) {
            throw new \RuntimeException("Could not open stream for {$this->xmlPath}");
        }

        $output = fopen($this->tempXmlFile, 'w');
        
        if ($output === false) {
            fclose($stream);
            throw new \RuntimeException("Could not create temp file");
        }

        // Copy chunk by chunck to keep as low as possible memory
        while (!feof($stream)) {
            $chunk = fread($stream, 8192); // Read 8kb 
            fwrite($output, $chunk);
        }

        fclose($stream);
        fclose($output);

        return $this->tempXmlFile;
    }

    /**
     * Iterates on all lines (streaming)
     */
    public function getRowIterator(): \Generator
    {
        $stream = $this->reader->getZip()->getStream($this->xmlPath);
        echo "Mémoire après récupération du stream " . round(memory_get_usage(true) / 1024 / 1024, 2) . " MB\n";
        if ($stream === false) {
            throw new \RuntimeException("Could not open stream for {$this->xmlPath}");
        }
        $buffer = '';
        $rowNumber = 0;

       while (!feof($stream)) {
            $chunk = fread($stream, 8192);
            $buffer .= $chunk;

            // Parser lines
            while (preg_match('/<row[^>]*>(.*?)<\/row>/s', $buffer, $match, PREG_OFFSET_CAPTURE)) {
                $rowNumber++;
                $rowXml = $match[0][0];
                $matchStart = $match[0][1];

                //Create the row with XML already parsed (pas besoin de stocker l'offset)
                yield $this->createRowFromXml($rowNumber, $rowXml);

                // Remove the line from the buffer
                $buffer = substr($buffer, $matchStart + strlen($rowXml));
            }

            // Keep a minimal buffer minimum to not split balises
            if (strlen($buffer) > 50000) {
                $buffer = substr($buffer, -10240);
            }
        }

        fclose($stream);
    }

    private function createRowFromXml(int $rowNumber, string $rowXml): Row
    {
        return new Row(
            sheet: $this,
            number: $rowNumber,
            rowXml: $rowXml
        );
    }

    /**
     * Build the index of row positions
     */
    public function buildIndex(): self
    {
        if ($this->indexed) {
            return $this;
        }

        $this->rowIndex = new RowIndex();
        $tempFile = $this->ensureTempFile();

        $handle = fopen($tempFile, 'r');
        if ($handle === false) {
            throw new \RuntimeException("Could not open temp file");
        }

        $buffer = '';
        $rowNumber = 0;
        $globalPosition = 0;

        while (!feof($handle)) {
            $chunk = fread($handle, 8192);
            $buffer .= $chunk;

            $offset = 0;
            while (preg_match('/<row[^>]*>.*?<\/row>/s', $buffer, $match, PREG_OFFSET_CAPTURE, $offset)) {
                $rowNumber++;
                $matchStart = $match[0][1];
                $rowLength = strlen($match[0][0]);

                $startOffset = $globalPosition + $matchStart;

                $this->rowIndex->add($rowNumber, $startOffset, $rowLength);

                $offset = $matchStart + $rowLength;
            }

            if ($offset > 0) {
                $globalPosition += $offset;
                $buffer = substr($buffer, $offset);
            }

            if (strlen($buffer) > 10000) {
                $keepSize = 5000;
                $globalPosition += strlen($buffer) - $keepSize;
                $buffer = substr($buffer, -$keepSize);
            }
        }

        fclose($handle);
        $this->indexed = true;

        return $this;
    }

    public function __destruct()
    {
        if ($this->tempXmlFile !== null && file_exists($this->tempXmlFile)) {
            @unlink($this->tempXmlFile);
        }
    }
}
