<?php

namespace Sigi\ExcelReader\Reader;

class SharedStringsReader
{
    private ?string $tempXmlFile = null;
    private array $index = []; // Position de chaque string
    private bool $indexed = false;
    private array $cache = []; // Cache LRU
    private array $cacheHits = []; // Tracking pour LRU
    private int $cacheMaxSize = 500;
    private ?string $sharedStringsPath = null;

    public function __construct(
        private \ZipArchive $zip,
        ?string $sharedStringsPath = null
    ) {
        $this->sharedStringsPath = $sharedStringsPath;
    }

    /**
     * Récupère une shared string par son index (lazy loading)
     */
    public function get(int $index): string
    {
        // Vérifier le cache
        if (isset($this->cache[$index])) {
            $this->cacheHits[$index] = microtime(true);
            return $this->cache[$index];
        }

        // Pas de shared strings dans ce fichier
        if ($this->sharedStringsPath === null) {
            return '';
        }

        if (!$this->indexed) {
            $this->buildIndex();
        }

        if (!isset($this->index[$index])) {
            return '';
        }

        $position = $this->index[$index];
        $value = $this->readStringAt($position);

        // Ajouter au cache avec éviction LRU si nécessaire
        if (count($this->cache) >= $this->cacheMaxSize) {
            $this->evictOldest();
        }

        $this->cache[$index] = $value;
        $this->cacheHits[$index] = microtime(true);

        return $value;
    }

    /**
     * Vérifie si des shared strings existent
     */
    public function hasSharedStrings(): bool
    {
        return $this->sharedStringsPath !== null;
    }

    /**
     * Retourne le nombre total de shared strings
     */
    public function count(): int
    {
        if (!$this->indexed) {
            $this->buildIndex();
        }

        return count($this->index);
    }

    /**
     * Vide le cache (libère la mémoire)
     */
    public function clearCache(): void
    {
        $this->cache = [];
        $this->cacheHits = [];
    }

    /**
     * Réduit le cache à N entrées (garde les plus récentes)
     */
    public function trimCache(int $keepCount = 100): void
    {
        if (count($this->cache) <= $keepCount) {
            return;
        }

        // Trier par timestamp (les plus récents en premier)
        arsort($this->cacheHits);
        $keysToKeep = array_slice(array_keys($this->cacheHits), 0, $keepCount, true);

        $newCache = [];
        $newHits = [];

        foreach ($keysToKeep as $key) {
            $newCache[$key] = $this->cache[$key];
            $newHits[$key] = $this->cacheHits[$key];
        }

        $this->cache = $newCache;
        $this->cacheHits = $newHits;
    }

    /**
     * Pré-charge les N premiers strings dans le cache
     * (utile si les premiers strings sont les plus utilisés)
     */
    public function warmupCache(int $count = 500): void
    {
        if (!$this->indexed) {
            $this->buildIndex();
        }

        $maxCount = min($count, count($this->index));

        for ($i = 0; $i < $maxCount; $i++) {
            $this->get($i);
        }
    }

    /**
     * Retourne l'utilisation mémoire
     */
    public function getMemoryUsage(): int
    {
        $indexSize = count($this->index) * 40; // ~40 bytes par entrée
        $cacheSize = 0;

        foreach ($this->cache as $string) {
            $cacheSize += strlen($string) + 100; // String + overhead PHP
        }

        return $indexSize + $cacheSize;
    }

    /**
     * Statistiques du cache
     */
    public function getCacheStats(): array
    {
        return [
            'size' => count($this->cache),
            'max_size' => $this->cacheMaxSize,
            'memory_bytes' => $this->getMemoryUsage(),
            'total_strings' => count($this->index),
        ];
    }

    /**
     * Construit un index des positions sans charger les valeurs
     */
    private function buildIndex(): void
    {
        $this->indexed = true;

        if ($this->sharedStringsPath === null) {
            return;
        }

        // Récupérer le contenu XML
        $xmlContent = $this->zip->getFromName($this->sharedStringsPath);

        if ($xmlContent === false) {
            return;
        }

        // Extraire temporairement pour permettre fseek
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

            // Chercher les balises <si> (shared item)
            $offset = 0;
            while (preg_match('/<si>.*?<\/si>/s', $buffer, $match, PREG_OFFSET_CAPTURE, $offset)) {
                $matchStart = $match[0][1];
                $matchLength = strlen($match[0][0]);

                // Stocker la position de ce string
                $this->index[$stringIndex] = [
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

            // Garder un buffer minimum
            if (strlen($buffer) > 10000) {
                $keepSize = 5000;
                $globalPosition += strlen($buffer) - $keepSize;
                $buffer = substr($buffer, -$keepSize);
            }
        }

        fclose($handle);
    }

    /**
     * Lit une string à une position donnée
     */
    private function readStringAt(array $position): string
    {
        if ($this->tempXmlFile === null || !file_exists($this->tempXmlFile)) {
            return '';
        }

        $handle = fopen($this->tempXmlFile, 'r');
        if ($handle === false) {
            return '';
        }

        fseek($handle, $position['start']);
        $xml = fread($handle, $position['length']);
        fclose($handle);

        return $this->extractTextFromSi($xml);
    }

    /**
     * Extrait le texte d'un élément <si>
     */
    private function extractTextFromSi(string $xml): string
    {
        // Cas 1 : Texte simple <si><t>Texte</t></si>
        if (preg_match('/<t[^>]*>([^<]*)<\/t>/', $xml, $match)) {
            return $this->decodeXmlEntities($match[1]);
        }

        // Cas 2 : Texte avec formatage <si><r><t>Texte</t></r></si>
        if (preg_match_all('/<t[^>]*>([^<]*)<\/t>/', $xml, $matches)) {
            $text = implode('', $matches[1]);
            return $this->decodeXmlEntities($text);
        }

        // Cas 3 : Texte riche avec plusieurs éléments <r>
        if (preg_match_all('/<r>.*?<t[^>]*>([^<]*)<\/t>.*?<\/r>/s', $xml, $matches)) {
            $text = implode('', $matches[1]);
            return $this->decodeXmlEntities($text);
        }

        return '';
    }

    /**
     * Décode les entités XML
     */
    private function decodeXmlEntities(string $text): string
    {
        $text = str_replace('&lt;', '<', $text);
        $text = str_replace('&gt;', '>', $text);
        $text = str_replace('&quot;', '"', $text);
        $text = str_replace('&apos;', "'", $text);
        $text = str_replace('&amp;', '&', $text);

        return $text;
    }

    /**
     * Éviction LRU (Least Recently Used)
     */
    private function evictOldest(): void
    {
        if (empty($this->cacheHits)) {
            return;
        }

        // Trouver la clé avec le timestamp le plus ancien
        asort($this->cacheHits);
        $oldestKey = array_key_first($this->cacheHits);

        unset($this->cache[$oldestKey]);
        unset($this->cacheHits[$oldestKey]);
    }

    /**
     * Nettoie les fichiers temporaires
     */
    public function __destruct()
    {
        if ($this->tempXmlFile !== null && file_exists($this->tempXmlFile)) {
            @unlink($this->tempXmlFile);
        }
    }
}
