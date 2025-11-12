<?php 

declare(strict_types=1);

namespace Sigi\ExcelReader\Tests\Loader;

use bovigo\vfs\vfsStream;
use bovigo\vfs\vfsStreamFile;
use PHPUnit\Framework\TestCase;
use bovigo\vfs\vfsStreamDirectory;
use Sigi\ExcelReader\Loader\SheetLoader;
use Sigi\ExcelReader\Exception\FileNotFoundException;
use Sigi\ExcelReader\Exception\FileNotReadableException;


class SheetLoaderTest extends TestCase
{

    private vfsStreamDirectory $root;
    private SheetLoader $loader;
    protected function setUp(): void
    {
        $this->root = vfsStream::setup();
        $this->loader = new SheetLoader();
    }

    public function testTriggerException(): void
    {
        $this->expectException(FileNotFoundException::class);
        $open = $this->loader->open("./unknownfile.xlsx");

        $notRedeableFile = new vfsStreamFile('file.xlsx', 0000);
        $this->root->addChild($notRedeableFile);

        $this->expectException(FileNotReadableException::class);
        $open = $this->loader->open($notRedeableFile->path());
    }

    public function testLoad(): void
    {
        $this->loader->open("xlsx_with_300k_rows_and_inline_strings.xlsx");
    }
}