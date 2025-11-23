# Xola ReportWriter Bundle

A lightweight Symfony bundle for generating tabular data exports (CSV and Excel) from arbitrary structured arrays. Supports nested headers (two-level) and automatic column sizing in Excel via PhpSpreadsheet.

## Requirements

- PHP: ^7.4
- Symfony Framework Bundle: ^5.4
- Extensions: `ext-json`
- Libraries: `phpoffice/phpspreadsheet`

## Installation

```bash
composer require xola/reportwriter-bundle
```

## Usage Overview

Register the bundle normally in your Symfony application's `bundles.php` (Symfony Flex usually auto-registers):

```php
return [
    Xola\ReportWriterBundle\XolaReportWriterBundle::class => ['all' => true],
];
```

### Services

The bundle exposes two writer services:

- `csvwriter` (`Xola\\ReportWriterBundle\\Service\\CSVWriter`) – writes CSV files.
- `excelwriter` (`Xola\\ReportWriterBundle\\Service\\ExcelWriter`) – writes XLSX files using PhpSpreadsheet.

Each writer implements the following workflow:

1. `setup($filepath)` – initialize output target.
2. `cacheData($filename, $data)` – append JSON-encoded rows to a temp cache file.
3. `parseHeaders($data, $existingHeaders)` – gather/merge headers (supports nested headers arrays).
4. `prepare($cacheFile, $sortedHeaders, $freezeHeaders = false)` (Excel) / `prepare($cacheFile, $sortedHeaders)` (CSV) – consume cache and write rows.
5. `finalize()` – flush/close underlying resources.

### Example

```php
$excelWriter = $container->get('excelwriter');
$excelWriter->setup('/tmp/report.xlsx');
$headers = $excelWriter->parseHeaders($firstRow); // or build manually
$excelWriter->cacheData($cacheFile, [$firstRow, $secondRow]);
$excelWriter->prepare($cacheFile, $headers, true); // freeze headers
$excelWriter->finalize();
```

### Nested Headers Format

Use an array structure like:

```php
$headers = [
    'Order ID',
    ['Customer' => ['Name', 'Email']],
    ['Totals' => ['Subtotal', 'Tax', 'Grand Total']],
];
```

### Separators (Excel Only)

Insert a row containing the string `ExcelWriter::SEPARATOR` into the cache file to produce a bottom border between logical sections.

## Testing

```bash
composer test
```

## Continuous Integration

A GitHub Actions workflow (`.github/workflows/ci.yml`) runs the test suite on PHP 7.4 and Symfony 5.4.

## Contributing

1. Fork & create feature branch.
2. Add/adjust tests.
3. Submit PR.

## License

MIT
