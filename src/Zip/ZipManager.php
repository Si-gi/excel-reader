<?php

declare(strict_types=1);

namespace Sigi\ExcelReader\Zip;

use ZipArchive;
use Sigi\ExcelReader\Exception\FileNotFoundException;

class ZipManager
{
    private ?ZipArchive $zip = null;
    private ?string $filePath = null;

    public function open(string $file): self
    {
        if (!file_exists($file)) {
            throw new FileNotFoundException("File not found: {$file}");
        }

        if (!is_readable($file)) {
            throw new \RuntimeException("File is not readable: {$file}");
        }

        $this->filePath = $file;
        $this->zip = new ZipArchive();

        if ($this->zip->open($file) !== true) {
            throw new \RuntimeException("Could not open ZIP archive: {$file}");
        }

        return $this;
    }

    public function close(): void
    {
        if ($this->zip !== null) {
            $this->zip->close();
            $this->zip = null;
        }
        $this->filePath = null;
    }

    public function getZip(): ZipArchive
    {
        if ($this->zip === null) {
            throw new \RuntimeException("No ZIP opened. Call open() first.");
        }

        return $this->zip;
    }

    public function getFilePath(): string
    {
        if ($this->filePath === null) {
            throw new \RuntimeException("No file opened.");
        }

        return $this->filePath;
    }

    public function isOpen(): bool
    {
        return $this->zip !== null;
    }

    public function __destruct()
    {
        $this->close();
    }
}