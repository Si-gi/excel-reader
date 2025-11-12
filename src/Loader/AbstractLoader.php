<?php 

declare(strict_types=1);

namespace Sigi\ExcelReader\Loader;

use ZipArchive;
use Sigi\ExcelReader\Exception\FileNotFoundException;

abstract class AbstractLoader {

   
    private ?string $filePath = null;
    public function __construct(private ZipArchive $zip)
    {

    }
    public function open(string $file): static
    {
        if (!file_exists($file)) {
            throw new FileNotFoundException("File not found: {$file}");
        }
        $this->filePath = $file;

        if ($this->zip->open($file) !== true) {
            throw new \RuntimeException("Could not open ZIP archive: {$file}");
        }
        return $this;
    }
}