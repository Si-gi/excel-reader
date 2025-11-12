<?php

declare(strict_types=1);

namespace Sigi\ExcelReader\Reader;

use ZipArchive;
use Sigi\ExcelReader\Objects\Sheet;
use Sigi\ExcelReader\Objects\SheetCollection;
use Sigi\ExcelReader\Exception\FileNotFoundException;
use Sigi\ExcelReader\Exception\FileNotReadableException;

class Reader
{
    private array $tempFiles = [];
    private ?ZipArchive $zip = null;
    private ?string $filePath = null;
    private ?SheetCollection $sheets = null;
    private ?SharedStringsReader $sharedStringsReader = null;
    private array $sharedStrings = [];

    public function getTempFiles(): array
    {
        return $this->tempFiles;
    }
    public function open(string $file): self
    {
        if (!file_exists($file)) {
            throw new FileNotFoundException("File not found: {$file}");
        }

        $this->filePath = $file;
        $this->zip = new ZipArchive();

        if ($this->zip->open($file) !== true) {
            throw new \RuntimeException("Could not open ZIP archive: {$file}");
        }
        $this->sharedStringsReader = new SharedStringsReader($this->zip, $file);
        // $this->loadSharedStrings();

        $this->loadSheets();

        return $this;
    }

    public function getSharedStringsReader(): SharedStringsReader
    {
        if ($this->sharedStringsReader === null) {
            throw new \RuntimeException("No file opened.");
        }

        return $this->sharedStringsReader;
    }

    public function close(): void
    {
        if ($this->zip !== null) {
            $this->zip->close();
            $this->zip = null;
        }

        $this->sheets = null;
        $this->sharedStrings = [];
        $this->sharedStringsReader = null;
        foreach ($this->tempFiles as $tempFile) {
            if (file_exists($tempFile)) {
                @unlink($tempFile);
            }
        }

        $this->tempFiles = [];
    }

    public function getSheets(): SheetCollection
    {
        if ($this->sheets === null) {
            throw new \RuntimeException("No file opened. Call open() first.");
        }

        return $this->sheets;
    }

    public function getSharedStrings(): array
    {
        return $this->sharedStrings;
    }

    public function getZip(): ZipArchive
    {
        if ($this->zip === null) {
            throw new \RuntimeException("No file opened.");
        }

        return $this->zip;
    }

    public function getFilePath(): string
    {
        return $this->filePath;
    }

    private function loadSharedStrings(): void
    {
        $xml = $this->zip->getFromName('xl/sharedStrings.xml');

        if ($xml === false) {
            return;
        }

        $reader = new \XMLReader();
        $reader->XML($xml);

        while ($reader->read()) {
            if ($reader->nodeType === \XMLReader::ELEMENT && str_contains($reader->name, 't')) {
                $this->sharedStrings[] = $reader->readString();
            }
        }

        $reader->close();
    }

    private function loadSheets(): void
    {
        $this->sheets = new SheetCollection();

        $workbookXml = $this->zip->getFromName('xl/workbook.xml');

        if ($workbookXml === false) {
            throw new \RuntimeException("Could not read workbook.xml");
        }

        $reader = new \XMLReader();
        $reader->XML($workbookXml);

        $sheetIndex = 0;

        while ($reader->read()) {
            if ($reader->nodeType === \XMLReader::ELEMENT && str_contains($reader->name,'sheet')) {
                $sheetName = $reader->getAttribute('name');
                $xmlPath = $this->getSheetXmlPath($sheetIndex);

                $sheet = new Sheet(
                    reader: $this,
                    name: $sheetName,
                    index: $sheetIndex,
                    xmlPath: $xmlPath
                );

                $this->sheets->add($sheet);
                $sheetIndex++;
            }
        }

        $reader->close();
    }

    private function getSheetXmlPath(int $index): string
    {
        $sheetNum = $index + 1;
        $path = "xl/worksheets/sheet{$sheetNum}.xml";

        if ($this->zip->locateName($path) !== false) {
            return $path;
        }

        // Chercher le bon fichier
        for ($i = 1; $i <= 100; $i++) {
            $testPath = "xl/worksheets/sheet{$i}.xml";
            if ($this->zip->locateName($testPath) !== false && ($i - 1) === $index) {
                return $testPath;
            }
        }

        throw new \RuntimeException("Could not find XML file for sheet index {$index}");
    }

    public function extractSheet(string $xmlPath): string
    {
        if (isset($this->tempFiles[$xmlPath])) {
            return $this->tempFiles[$xmlPath];
        }

        $tempFile = sys_get_temp_dir() . '/excel_' . md5($xmlPath) . '.xml';
        $source = "zip://{$this->filePath}#{$xmlPath}";

        if (!copy($source, $tempFile)) {
            throw new \RuntimeException("Could not extract $xmlPath");
        }

        $this->tempFiles[$xmlPath] = $tempFile;
        return $tempFile;
    }
}