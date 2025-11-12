<?php

declare(strict_types=1);

namespace Sigi\ExcelReader\tests;

use bovigo\vfs\vfsStream;
use bovigo\vfs\vfsStreamFile;
use PHPUnit\Framework\TestCase;
use bovigo\vfs\vfsStreamDirectory;
use Sigi\ExcelReader\Reader;
use Sigi\ExcelReader\Exception\FileNotFoundException;
use Sigi\ExcelReader\Exception\FileNotReadableException;

class ReaderTest extends TestCase
{
    private Reader $reader;
    private vfsStreamDirectory $root;
    protected function setUp(): void
    {

        $this->root = vfsStream::setup();
        $this->reader = new Reader();
    }

    public function testTriggerException(): void
    {
        $this->expectException(FileNotFoundException::class);
        $open = $this->reader->open("./unknownfile.xlsx");

        $notRedeableFile = new vfsStreamFile('file.xlsx', 0000);
        $this->root->addChild($notRedeableFile);

        $this->expectException(FileNotReadableException::class);
        $open = $this->reader->open($notRedeableFile->path());
    }
    public function testOpen(): void
    {
        $file = new vfsStreamFile('file.xlsx');
        $this->root->addChild($file);
        $filePath = vfsStream::url($file->path());

        $open = $this->reader->open($filePath);

        $this->assertInstanceOf(Reader::class, $open);
    }

    public function testGetWorkSheet(): void
    {
        $this->reader->open($filePath);
        
    }
}