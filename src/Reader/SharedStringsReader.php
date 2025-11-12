<?php

namespace Sigi\ExcelReader\Reader;

class SharedStringsReader
{
    private ?string $tempXmlFile = null;
    private array $stringIndex = [];
    private bool $indexed = false;
    private array $cache = []; // TODO: delete
    private int $cacheMaxSize = 500;
    private array $cacheHits = [];

    public function __construct(private \ZipArchive $zip, private string $filePath)
    {
    }

    public function getCacheMaxSize(): int
    {
        return $this->cacheMaxSize;
    }

    public function setCacehMaxSize(int $cacheMaxSize): self
    {
        $this->cacheMaxSize = $cacheMaxSize;

        return $this;
    }

    public function clearCache(): void
    {
        $this->cache = [];
        $this->cacheHits = [];
    }

    /**
     * Get sharedString by index (lazy loading)
     */
    public function get(int $index): string
    {
        if (isset($this->cache[$index])) {
            // $this->cacheHits[$index] = time();
            return $this->cache[$index];
        }

        if (!$this->indexed) {
            $this->buildIndex();
        }
        if (!isset($this->stringIndex[$index])) {
            return '';
        }

        $position = $this->stringIndex[$index];
        $value = $this->readStringAt($position);

        $this->cache[$index] = $value;
        // $this->cacheHits[$index] = time();
        return $value;
    }

    /**
     * Build an index of positions without loading values
     */
    private function buildIndex(): void
    {
        $this->indexed = true;

        // Check if file exist
        $xmlContent = $this->zip->getFromName('xl/sharedStrings.xml');
        
        if ($xmlContent === false) {
            // no shared strign
            return;
        }

        // Extract to allow fopen
        $this->tempXmlFile = sys_get_temp_dir() . '/excel_shared_strings_' . uniqid() . '.xml';
        file_put_contents($this->tempXmlFile, $xmlContent);
        $handle = fopen($this->tempXmlFile, 'r');
        if ($handle === false) {
            throw new \RuntimeException("Could not open shared strings temp file");
        }

        $buffer = '';
        $stringIndex = 0;
        $globalPosition = 0;

        while (!feof($handle)) {

            $chunk = fread($handle, 8192);
            $buffer .= $chunk;

            // search <si> (shared item)
            $offset = 0;
            while (preg_match('/<si>.*?<\/si>/s', $buffer, $match, PREG_OFFSET_CAPTURE, $offset)) {
                $matchStart = $match[0][1];
                $matchLength = strlen($match[0][0]);
                
                // Stock the index position
                $this->stringIndex[$stringIndex] = [
                    'start' => $globalPosition + $matchStart,
                    'length' => $matchLength
                ];

                $stringIndex++;
                $offset = $matchStart + $matchLength;
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
    }

    /**
     * Read a string at a given position
     */
    private function readStringAt(array $position): string
    {
        if ($this->tempXmlFile === null) {
            return '';
        }

        $handle = fopen($this->tempXmlFile, 'r');
        if ($handle === false) {
            return '';
        }

        fseek($handle, $position['start']);
        $xml = fread($handle, $position['length']);
        fclose($handle);

        if (preg_match('/<t[^>]*>([^<]*)<\/t>/', $xml, $match)) {
            return $match[1];
        }

        if (preg_match_all('/<t[^>]*>([^<]*)<\/t>/', $xml, $matches)) {
            return implode('', $matches[1]);
        }

        return '';
    }

    /**
     * Garde seulement les N entrées les plus récentes
     */
    public function trimCache(int $keepCount = 100): void
    {
        if (count($this->cache) <= $keepCount) {
            return;
        }

        arsort($this->cacheHits);
        $keysToKeep = array_slice(array_keys($this->cacheHits), 0, $keepCount);

        $newCache = [];
        foreach ($keysToKeep as $key) {
            $newCache[$key] = $this->cache[$key];
        }

        $this->cache = $newCache;
        $this->cacheHits = array_intersect_key($this->cacheHits, $newCache);
    }

    /**
     * Éviction LRU (Least Recently Used)
     */
    private function evictOldest(): void
    {
        if (empty($this->cacheHits)) {
            return;
        }

        asort($this->cacheHits);
        $oldestKey = array_key_first($this->cacheHits);

        unset($this->cache[$oldestKey]);
        unset($this->cacheHits[$oldestKey]);
    }

    public function getMemoryUsage(): int
    {
        $indexSize = count($this->stringIndex) * 40; // ~40 bytes par entrée
        $cacheSize = 0;
        
        foreach ($this->cache as $string) {
            $cacheSize += strlen($string) + 100; // String + overhead
        }

        return $indexSize + $cacheSize;
    }

    public function __destruct()
    {
        if ($this->tempXmlFile !== null && file_exists($this->tempXmlFile)) {
            @unlink($this->tempXmlFile);
        }
    }
}