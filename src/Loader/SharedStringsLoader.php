<?php
declare(strict_types=1);

namespace Sigi\ExcelReader\Loader;

use Sigi\ExcelReader\Loader\AbstractLoader;
use Sigi\ExcelReader\Reader\SharedStringsReader;

class SharedStringsLoader extends AbstractLoader
{
    private ?SharedStringsReader $sharedStringsReader = null;

    public function load(string $fromName = ''): static
    {
        if ($this->loaded) {
            return $this;
        }

        // Utiliser le chemin détecté par la structure
        $path = $fromName ?: $this->structure->sharedStringsPath;

        if (empty($path) || !$this->structure->hasSharedStrings) {
            // Pas de shared strings dans ce fichier
            $this->sharedStringsReader = new SharedStringsReader($this->zip, null);
            $this->loaded = true;
            return $this;
        }

        $this->sharedStringsReader = new SharedStringsReader($this->zip, $path);
        $this->loaded = true;

        return $this;
    }

    public function getSharedStringsReader(): SharedStringsReader
    {
        if (!$this->loaded) {
            throw new \RuntimeException("Shared strings not loaded. Call load() first.");
        }

        return $this->sharedStringsReader;
    }
}