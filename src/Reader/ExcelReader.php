<?php

namespace Sigi\ExcelReader\Reader;

use Sigi\ExcelReader\Zip\ZipManager;
use Sigi\ExcelReader\Loader\SheetLoader;
use Sigi\ExcelReader\Objects\SheetCollection;
use Sigi\ExcelReader\Detector\StructureDetector;
use Sigi\ExcelReader\Loader\SharedStringsLoader;

class ExcelReader
{
    private ZipManager $zipManager;
    private ?SheetLoader $sheetLoader = null;

    public function __construct()
    {
        $this->zipManager = new ZipManager();
    }

    public function open(string $file): self
    {
        $this->zipManager->open($file);

        $zip = $this->zipManager->getZip();

        $detector = new StructureDetector($zip);
        $structure = $detector->detect();

        $sharedStringsLoader = new SharedStringsLoader($zip, $structure);
        $sharedStringsLoader->load();

        $this->sheetLoader = new SheetLoader($zip, $structure);
        $this->sheetLoader
            ->setSharedStringsReader($sharedStringsLoader->getSharedStringsReader())
            ->load();

        return $this;
    }

    public function getSheets(): SheetCollection
    {
        if ($this->sheetLoader === null) {
            throw new \RuntimeException("No file opened.");
        }

        return $this->sheetLoader->getSheets();
    }

    public function close(): void
    {
        $this->zipManager->close();
    }
}