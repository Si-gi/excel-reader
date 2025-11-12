<?php

declare(strict_types=1);

namespace Sigi\ExcelReader\tests\Objects;

use PHPUnit\Framework\TestCase;

class WorksheetTest extends TestCase
{

    public function testGetCollection(): void
    {
        $collection = $this->worksheet->getSheetCollection();
        $this->assertEmpty($collection);
    }
}