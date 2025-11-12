<?php

declare(strict_types=1);

namespace Sigi\ExcelReader\Tests\Detector;

use PHPUnit\Framework\TestCase;
use Sigi\ExcelReader\Detector\ExcelStructure;

final class ExcelStructureTest extends TestCase
{
    private function makeStructure(array $overrides = []): ExcelStructure
    {
        $defaults = [
            'worksheetsPath' => 'xl/worksheets',
            'workbookPath' => 'xl/workbook.xml',
            'sharedStringsPath' => 'xl/sharedStrings.xml',
            'relationsPath' => 'xl/_rels/workbook.xml.rels',
            'namespaces' => ['http://schemas.openxmlformats.org/spreadsheetml/2006/main'],
            'availableSheets' => [0 => 'xl/worksheets/sheet1.xml', 1 => 'xl/worksheets/sheet2.xml'],
            'sheetElementName' => 'sheet',
            'relationshipIdAttribute' => 'r:id',
            'hasSharedStrings' => true,
            'metadata' => ['creator' => 'tester']
        ];

        $args = array_replace($defaults, $overrides);

        return new ExcelStructure(
            worksheetsPath: $args['worksheetsPath'],
            workbookPath: $args['workbookPath'],
            sharedStringsPath: $args['sharedStringsPath'],
            relationsPath: $args['relationsPath'],
            namespaces: $args['namespaces'],
            availableSheets: $args['availableSheets'],
            sheetElementName: $args['sheetElementName'],
            relationshipIdAttribute: $args['relationshipIdAttribute'],
            hasSharedStrings: $args['hasSharedStrings'],
            metadata: $args['metadata']
        );
    }

    public function testGetWorksheetPathReturnsPathForExistingIndex(): void
    {
        $structure = $this->makeStructure();
        self::assertSame('xl/worksheets/sheet1.xml', $structure->getWorksheetPath(0));
        self::assertSame('xl/worksheets/sheet2.xml', $structure->getWorksheetPath(1));
    }

    public function testGetWorksheetPathReturnsNullForMissingIndex(): void
    {
        $structure = $this->makeStructure();
        self::assertNull($structure->getWorksheetPath(42));
    }

    public function testSupportsNamespaceStrictComparison(): void
    {
        $structure = $this->makeStructure([
            'namespaces' => ['urn:a', 'urn:b']
        ]);

        self::assertTrue($structure->supportsNamespace('urn:a'));
        self::assertFalse($structure->supportsNamespace('URN:A'));
    }

    public function testPublicReadonlyPropertiesExposeConstructorValues(): void
    {
        $args = [
            'worksheetsPath' => 'p/worksheets',
            'workbookPath' => 'p/workbook.xml',
            'sharedStringsPath' => 'p/ss.xml',
            'relationsPath' => 'p/rels.xml.rels',
            'namespaces' => ['urn:x'],
            'availableSheets' => [5 => 'p/worksheets/sheet6.xml'],
            'sheetElementName' => 'worksheet',
            'relationshipIdAttribute' => 'rel:id',
            'hasSharedStrings' => false,
            'metadata' => ['k' => 'v']
        ];

        $structure = $this->makeStructure($args);

        self::assertSame($args['worksheetsPath'], $structure->worksheetsPath);
        self::assertSame($args['workbookPath'], $structure->workbookPath);
        self::assertSame($args['sharedStringsPath'], $structure->sharedStringsPath);
        self::assertSame($args['relationsPath'], $structure->relationsPath);
        self::assertSame($args['namespaces'], $structure->namespaces);
        self::assertSame($args['availableSheets'], $structure->availableSheets);
        self::assertSame($args['sheetElementName'], $structure->sheetElementName);
        self::assertSame($args['relationshipIdAttribute'], $structure->relationshipIdAttribute);
        self::assertSame($args['hasSharedStrings'], $structure->hasSharedStrings);
        self::assertSame($args['metadata'], $structure->metadata);
    }

    public function testGetWorksheetPathWorksWithSparseIndices(): void
    {
        $structure = $this->makeStructure([
            'availableSheets' => [2 => 'xl/worksheets/sheet3.xml']
        ]);

        self::assertNull($structure->getWorksheetPath(0));
        self::assertSame('xl/worksheets/sheet3.xml', $structure->getWorksheetPath(2));
    }
}
