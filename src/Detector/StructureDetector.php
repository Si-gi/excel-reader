<?php

namespace Sigi\ExcelReader\Detector;

use ZipArchive;

class StructureDetector
{
    private ZipArchive $zip;
    private array $fileList = [];

    public function __construct(ZipArchive $zip)
    {
        $this->zip = $zip;
        $this->buildFileList();
    }

    /**
     * Détecte la structure complète du fichier Excel
     */
    public function detect(): ExcelStructure
    {
        return new ExcelStructure(
            worksheetsPath: $this->detectWorksheetsPath(),
            workbookPath: $this->detectWorkbookPath(),
            sharedStringsPath: $this->detectSharedStringsPath(),
            relationsPath: $this->detectRelationsPath(),
            namespaces: $this->detectNamespaces(),
            availableSheets: $this->detectAvailableSheets(),
            sheetElementName: $this->detectSheetElementName(),
            relationshipIdAttribute: $this->detectRelationshipIdAttribute(),
            hasSharedStrings: $this->hasSharedStrings(),
            metadata: $this->extractMetadata()
        );
    }

    /**
     * Construit la liste de tous les fichiers du ZIP
     */
    private function buildFileList(): void
    {
        $this->fileList = [];
        
        for ($i = 0; $i < $this->zip->numFiles; $i++) {
            $filename = $this->zip->getNameIndex($i);
            $this->fileList[] = $filename;
        }
    }

    /**
     * Détecte le chemin du dossier worksheets
     */
    private function detectWorksheetsPath(): string
    {
        $patterns = [
            'xl/worksheets/',
            'worksheets/',
            'Worksheets/',
            'xl/Worksheets/',
        ];

        foreach ($patterns as $pattern) {
            foreach ($this->fileList as $file) {
                if (str_starts_with($file, $pattern) && str_ends_with($file, '.xml')) {
                    return $pattern;
                }
            }
        }

        return 'xl/worksheets/'; // Défaut
    }

    /**
     * Détecte le chemin du workbook
     */
    private function detectWorkbookPath(): string
    {
        $patterns = [
            'xl/workbook.xml',
            'workbook.xml',
            'xl/Workbook.xml',
        ];

        foreach ($patterns as $pattern) {
            if ($this->zip->locateName($pattern) !== false) {
                return $pattern;
            }
        }

        throw new \RuntimeException("Could not find workbook.xml");
    }

    /**
     * Détecte le chemin des shared strings
     */
    private function detectSharedStringsPath(): string
    {
        $patterns = [
            'xl/sharedStrings.xml',
            'sharedStrings.xml',
            'xl/SharedStrings.xml',
        ];

        foreach ($patterns as $pattern) {
            if ($this->zip->locateName($pattern) !== false) {
                return $pattern;
            }
        }

        return ''; // Pas de shared strings
    }

    /**
     * Détecte le chemin des relations
     */
    private function detectRelationsPath(): string
    {
        $patterns = [
            'xl/_rels/workbook.xml.rels',
            '_rels/workbook.xml.rels',
            'xl/_rels/Workbook.xml.rels',
        ];

        foreach ($patterns as $pattern) {
            if ($this->zip->locateName($pattern) !== false) {
                return $pattern;
            }
        }

        return '';
    }

    /**
     * Détecte tous les namespaces utilisés
     */
    private function detectNamespaces(): array
    {
        $namespaces = [];
        $workbookXml = $this->zip->getFromName($this->detectWorkbookPath());

        if ($workbookXml === false) {
            return [];
        }

        // Extraire tous les xmlns
        preg_match_all('/xmlns:([^=]+)="([^"]+)"/', $workbookXml, $matches);

        for ($i = 0; $i < count($matches[1]); $i++) {
            $namespaces[$matches[1][$i]] = $matches[2][$i];
        }

        return $namespaces;
    }

    /**
     * Liste tous les fichiers worksheet disponibles
     */
    private function detectAvailableSheets(): array
    {
        $sheets = [];
        $worksheetsPath = $this->detectWorksheetsPath();

        foreach ($this->fileList as $file) {
            if (str_starts_with($file, $worksheetsPath) && 
                preg_match('/sheet\d+\.xml$/i', $file)) {
                $sheets[] = $file;
            }
        }

        // Trier par numéro
        usort($sheets, function($a, $b) {
            preg_match('/sheet(\d+)\.xml$/i', $a, $matchA);
            preg_match('/sheet(\d+)\.xml$/i', $b, $matchB);
            
            $numA = isset($matchA[1]) ? (int)$matchA[1] : 0;
            $numB = isset($matchB[1]) ? (int)$matchB[1] : 0;
            
            return $numA - $numB;
        });

        return $sheets;
    }

    /**
     * Détecte le nom de l'élément sheet (avec ou sans namespace)
     */
    private function detectSheetElementName(): string
    {
        $workbookXml = $this->zip->getFromName($this->detectWorkbookPath());

        if ($workbookXml === false) {
            return 'sheet';
        }

        $reader = new \XMLReader();
        $reader->XML($workbookXml);

        while ($reader->read()) {
            if ($reader->nodeType === \XMLReader::ELEMENT) {
                $name = strtolower($reader->name);
                $localName = strtolower($reader->localName);

                // Chercher un élément qui ressemble à "sheet"
                if ($localName === 'sheet' || str_ends_with($name, ':sheet')) {
                    $reader->close();
                    return $reader->name; // Retourner le nom complet avec préfixe
                }
            }
        }

        $reader->close();
        return 'sheet';
    }

    /**
     * Détecte le nom de l'attribut pour l'ID de relation
     */
    private function detectRelationshipIdAttribute(): string
    {
        $workbookXml = $this->zip->getFromName($this->detectWorkbookPath());

        if ($workbookXml === false) {
            return 'r:id';
        }

        $reader = new \XMLReader();
        $reader->XML($workbookXml);

        while ($reader->read()) {
            if ($reader->nodeType === \XMLReader::ELEMENT && 
                strtolower($reader->localName) === 'sheet') {
                
                // Chercher un attribut qui contient "id"
                if ($reader->hasAttributes) {
                    while ($reader->moveToNextAttribute()) {
                        $attrName = $reader->name;
                        if (str_contains(strtolower($attrName), 'id') && 
                            str_starts_with($reader->value, 'rId')) {
                            $reader->close();
                            return $attrName;
                        }
                    }
                }
                break;
            }
        }

        $reader->close();
        return 'r:id'; // Défaut
    }

    /**
     * Vérifie si le fichier contient des shared strings
     */
    private function hasSharedStrings(): bool
    {
        return !empty($this->detectSharedStringsPath());
    }

    /**
     * Extrait des métadonnées supplémentaires
     */
    private function extractMetadata(): array
    {
        $metadata = [
            'total_files' => count($this->fileList),
            'has_vba_macros' => $this->hasMacros(),
            'excel_version' => $this->detectExcelVersion(),
            'compression_ratio' => $this->calculateCompressionRatio(),
        ];

        return $metadata;
    }

    private function hasMacros(): bool
    {
        foreach ($this->fileList as $file) {
            if (str_contains($file, 'vbaProject')) {
                return true;
            }
        }
        return false;
    }

    private function detectExcelVersion(): string
    {
        // Analyser le content types pour détecter la version
        $contentTypes = $this->zip->getFromName('[Content_Types].xml');
        
        if ($contentTypes === false) {
            return 'unknown';
        }

        if (str_contains($contentTypes, 'officedocument.spreadsheetml')) {
            return 'Office Open XML (2007+)';
        }

        return 'unknown';
    }

    private function calculateCompressionRatio(): float
    {
        $compressedSize = 0;
        $uncompressedSize = 0;

        for ($i = 0; $i < $this->zip->numFiles; $i++) {
            $stat = $this->zip->statIndex($i);
            $compressedSize += $stat['comp_size'];
            $uncompressedSize += $stat['size'];
        }

        if ($uncompressedSize === 0) {
            return 0.0;
        }

        return round((1 - ($compressedSize / $uncompressedSize)) * 100, 2);
    }
}