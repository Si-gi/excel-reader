<?php 

declare(strict_types=1);
namespace Sigi\ExcelReader\Loader;

use PhpParser\Node\Name;


interface LoaderInterface {
    public function load(string $fromName): static;
}