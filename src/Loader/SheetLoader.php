<?php

declare(strict_types=1);

namespace Sigi\ExcelReader\Loader;

use Sigi\ExcelReader\Objects\Sheet;
use Sigi\ExcelReader\Loader\AbstractLoader;
use Sigi\ExcelReader\Objects\SheetCollection;
use Sigi\ExcelReader\Reader\SharedStringsReader;

class SheetLoader extends AbstractLoader {

    private ?SheetCollection $sheets = null;
    private ?SharedStringsReader $sharedStringsReader = null;

    public function setSharedStringsReader(SharedStringsReader $reader): self
    {
        $this->sharedStringsReader = $reader;
        return $this;
    }

    public function load(string $fromName = ''): static
    {
        if ($this->loaded) {
            return $this;
        }

        // Utiliser le chemin détecté par la structure
        $workbookPath = $fromName ?: $this->structure->workbookPath;

        $this->sheets = new SheetCollection();
        $workbookXml = $this->getXmlContent($workbookPath);

        $reader = new \XMLReader();
        $reader->XML($workbookXml);

        $sheetIndex = 0;
        $relations = $this->loadRelations();

        while ($reader->read()) {
            if ($this->isSheetElement($reader)) {
                $sheetName = $reader->getAttribute('name');
                $rId = $this->extractRelationId($reader);

                $xmlPath = $this->resolveSheetPath($rId, $sheetIndex, $relations);

                if ($xmlPath === null) {
                    throw new \RuntimeException("Could not find XML for sheet '{$sheetName}'");
                }

                $sheet = new Sheet(
                    zip: $this->zip,
                    structure: $this->structure,
                    sharedStringsReader: $this->sharedStringsReader,
                    name: $sheetName,
                    index: $sheetIndex,
                    xmlPath: $xmlPath
                );

                $this->sheets->add($sheet);
                $sheetIndex++;
            }
        }

        $reader->close();
        $this->loaded = true;

        return $this;
    }

    public function getSheets(): SheetCollection
    {
        if (!$this->loaded) {
            throw new \RuntimeException("Sheets not loaded. Call load() first.");
        }

        return $this->sheets;
    }

    /**
     * Vérifie si l'élément XML est une sheet (utilise la structure détectée)
     */
    private function isSheetElement(\XMLReader $reader): bool
    {
        if ($reader->nodeType !== \XMLReader::ELEMENT) {
            return false;
        }

        // Utiliser le nom d'élément détecté par la structure
        $expectedName = strtolower($this->structure->sheetElementName);
        $currentName = strtolower($reader->name);
        $localName = strtolower($reader->localName);

        return $localName === 'sheet' || 
               $currentName === $expectedName || 
               str_ends_with($currentName, ':sheet');
    }

    /**
     * Extrait l'ID de relation (utilise la structure détectée)
     */
    private function extractRelationId(\XMLReader $reader): ?string
    {
        // Essayer l'attribut détecté par la structure
        $rId = $reader->getAttribute($this->structure->relationshipIdAttribute);

        if (!empty($rId)) {
            return $rId;
        }

        // Fallback : essayer sans namespace
        $rId = $reader->getAttribute('id');

        if (!empty($rId)) {
            return $rId;
        }

        // Dernier recours : chercher n'importe quel attribut contenant "id"
        if ($reader->hasAttributes) {
            while ($reader->moveToNextAttribute()) {
                if (str_contains(strtolower($reader->name), 'id') && 
                    str_starts_with($reader->value, 'rId')) {
                    $value = $reader->value;
                    $reader->moveToElement();
                    return $value;
                }
            }
            $reader->moveToElement();
        }

        return null;
    }

    /**
     * Résout le chemin XML d'une sheet
     */
    private function resolveSheetPath(?string $rId, int $index, array $relations): ?string
    {
        var_dump($rId);
        var_dump($index);
        var_dump($relations);

        // 1. Essayer via les relations
        if (!empty($rId) && isset($relations[$rId])) {
            return $relations[$rId];
        }

        // 2. Essayer via la structure détectée
        return $this->structure->getWorksheetPath($index);
    }

    /**
     * Charge les relations workbook -> worksheets
     */
    private function loadRelations(): array
    {
        $relations = [];

        if (empty($this->structure->relationsPath)) {
            return $relations;
        }

        try {
            $relsXml = $this->getXmlContent($this->structure->relationsPath);
        } catch (\RuntimeException $e) {
            return $relations;
        }

        $reader = new \XMLReader();
        $reader->XML($relsXml);

        while ($reader->read()) {
            if ($reader->nodeType === \XMLReader::ELEMENT && 
                strtolower($reader->localName) === 'relationship') {
                
                $id = $reader->getAttribute('Id');
                $target = $reader->getAttribute('Target');
                $type = $reader->getAttribute('Type');

                if (!empty($id) && !empty($target) && 
                    (str_contains($type, 'worksheet') || str_contains($target, 'sheet'))) {
                    
                    $normalizedTarget = $this->normalizeTargetPath($target);
                    $relations[$id] = $normalizedTarget;
                }
            }
        }

        $reader->close();

        return $relations;
    }

    /**
     * Normalise un chemin de target
     */
    private function normalizeTargetPath(string $target): string
    {
        // Retirer les ../ relatifs
        $normalized = str_replace('../', '', $target);
        
        // Ajouter xl/ si nécessaire
        if (!str_starts_with($normalized, 'xl/') && 
            !str_starts_with($normalized, 'worksheets/')) {
            $normalized = 'xl/' . $normalized;
        }

        return $normalized;
    }
}