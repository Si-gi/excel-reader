<?php

namespace Sigi\ExcelReader\Detector;

class ExcelStructure
{
    public function __construct(
        public readonly string $worksheetsPath,
        public readonly string $workbookPath,
        public readonly string $sharedStringsPath,
        public readonly string $relationsPath,
        public readonly array $namespaces,
        public readonly array $availableSheets,
        public readonly string $sheetElementName,
        public readonly string $relationshipIdAttribute,
        public readonly bool $hasSharedStrings,
        public readonly array $metadata = []
    ) {
    }

    public function getWorksheetPath(int $index): ?string
    {
        return $this->availableSheets[$index] ?? null;
    }

    public function supportsNamespace(string $namespace): bool
    {
        return in_array($namespace, $this->namespaces, true);
    }
}