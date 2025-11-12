<?php 

declare(strict_types=1);

namespace Sigi\ExcelReader\Loader;

use ZipArchive;
use Sigi\ExcelReader\Detector\ExcelStructure;
use Sigi\ExcelReader\Exception\FileNotFoundException;

abstract class AbstractLoader implements LoaderInterface
{
    protected bool $loaded = false;

    public function __construct(
        protected ZipArchive $zip,
        protected ExcelStructure $structure
    ) {
    }

    public function isLoaded(): bool
    {
        return $this->loaded;
    }

    protected function getXmlContent(string $path): string
    {
        $content = $this->zip->getFromName($path);

        if ($content === false) {
            throw new \RuntimeException("Could not read file: {$path}");
        }

        return $content;
    }

    protected function getXmlStream(string $path)
    {
        $stream = $this->zip->getStream($path);

        if ($stream === false) {
            throw new \RuntimeException("Could not open stream for: {$path}");
        }

        return $stream;
    }

}
