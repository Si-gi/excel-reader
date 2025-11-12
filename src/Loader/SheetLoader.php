<?php

declare(strict_types=1);

namespace Sigi\ExcelReader\Loader;

use Sigi\ExcelReader\Objects\Sheet;
use Sigi\ExcelReader\Objects\SheetCollection;

class SheetLoader extends AbstractLoader implements LoaderInterface
{

    public function __construct()
    {
        parent::__construct();
    }
    public function load(string $fromName): static
    {
        $sheetCollection = new SheetCollection();
        $workbookXml = $this->zip->getFromName($fromName);

        if ($workbookXml === false) {
            throw new \RuntimeException("Could not read workbook.xml");
        }
        $reader = new \XMLReader();
        $reader->XML($workbookXml);

        $sheetIndex = 0;
        $sheetRelations = $this->getSheetRelations();

        while ($reader->read()) {
            if ($reader->nodeType === \XMLReader::ELEMENT && 
                (strtolower($reader->localName) === 'sheet' || str_contains(strtolower($reader->name), 'sheet'))) {
                
                $sheetName = $reader->getAttribute('name');
                $rId = $reader->getAttribute('r:id'); // TODO : maange by excel file

                // Essayer différentes façons d'obtenir le r:id
                if (empty($rId)) {
                    $rId = $reader->getAttributeNs('id', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships');
                }

                // Trouver le chemin XML via les relations ou par index
                $xmlPath = null;
                
                if (!empty($rId) && isset($sheetRelations[$rId])) {
                    $xmlPath = $sheetRelations[$rId];
                } else {
                    $xmlPath = $this->findSheetXmlByIndex($sheetIndex);
                }

                if ($xmlPath === null) {
                    throw new \RuntimeException("Could not find XML path for sheet '{$sheetName}' (index: {$sheetIndex})");
                }

                $sheet = new Sheet(
                    reader: $this,
                    name: $sheetName,
                    index: $sheetIndex,
                    xmlPath: $xmlPath
                );

                $sheetCollection->add($sheet);
                $sheetIndex++;
            }
        }

        $reader->close();
        return $this;
    }

    /**
     * Read the workbook.xml.rels to get relations
     */
    private function getSheetRelations(): array
    {
        $relations = [];

        $relsXml = $this->zip->getFromName('xl/_rels/workbook.xml.rels');

        if ($relsXml === false) {
            return $relations;
        }

        $reader = new \XMLReader();
        $reader->XML($relsXml);

        while ($reader->read()) {
            if ($reader->nodeType === \XMLReader::ELEMENT && 
                (strtolower($reader->localName) === 'relationship' || strtolower($reader->name) === 'relationship')) {
                
                $id = $reader->getAttribute('Id');
                $target = $reader->getAttribute('Target');
                $type = $reader->getAttribute('Type');

                // Vérifier que c'est bien une relation vers une worksheet
                if (!empty($id) && !empty($target) && 
                    (str_contains($type, 'worksheet') || str_contains($target, 'worksheet'))) {
                    
                    // Normaliser le chemin (enlever ../ si présent)
                    $normalizedTarget = str_replace('../', '', $target);
                    
                    // Ajouter xl/ si ce n'est pas déjà là
                    if (!str_starts_with($normalizedTarget, 'xl/')) {
                        $normalizedTarget = 'xl/' . $normalizedTarget;
                    }

                    $relations[$id] = $normalizedTarget;
                }
            }
        }

        $reader->close();

        return $relations;
    }

    /**
     * Find XML file of sheet by index (fallback)
     */
    private function findSheetXmlByIndex(int $index): ?string
    {
        // TODO: manage sheets paths
        $possiblePaths = [
            "xl/worksheets/sheet" . ($index + 1) . ".xml",
            "worksheets/sheet" . ($index + 1) . ".xml",
        ];

        // Chercher le premier qui existe
        foreach ($possiblePaths as $path) {
            if ($this->zip->locateName($path) !== false) {
                return $path;
            }
        }

        // Si rien trouvé, lister tous les fichiers worksheet disponibles
        $worksheets = $this->listWorksheetFiles();

        if (isset($worksheets[$index])) {
            return $worksheets[$index];
        }

        return null;
    }

    /**
     * Liste tous les fichiers worksheet dans le ZIP
     */
    private function listWorksheetFiles(): array
    {
        $worksheets = [];

        for ($i = 0; $i < $this->zip->numFiles; $i++) {
            $filename = $this->zip->getNameIndex($i);
            
            if (preg_match('/worksheets\/sheet\d+\.xml$/', $filename)) {
                $worksheets[] = $filename;
            }
        }

        // Trier par numéro de sheet
        usort($worksheets, function($a, $b) {
            preg_match('/sheet(\d+)\.xml$/', $a, $matchA);
            preg_match('/sheet(\d+)\.xml$/', $b, $matchB);
            
            $numA = isset($matchA[1]) ? (int)$matchA[1] : 0;
            $numB = isset($matchB[1]) ? (int)$matchB[1] : 0;
            
            return $numA - $numB;
        });

        return $worksheets;
    }
}