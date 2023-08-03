<?php

namespace Xola\ReportWriterBundle\Service;

use PhpOffice\PhpSpreadsheet\Shared\Font;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Worksheet\PageSetup;
use Psr\Log\LoggerInterface;
use Xola\ReportWriterBundle\SpreadsheetFactory;

class ExcelWriter extends AbstractWriter
{
    /** @var SpreadsheetFactory */
    private $phpexcel;
    /* @var Spreadsheet $spreadsheet */
    private $spreadsheet;
    private $currentRow = 1;

    /**
     * Separates sections of the report.
     * The spreadsheet will have a bottom border on the row preceding this separator.
     */
    const SEPARATOR = '---';

    public function __construct(LoggerInterface $logger, SpreadsheetFactory $phpExcel)
    {
        parent::__construct($logger);
        $this->phpexcel = $phpExcel;
        $this->spreadsheet = $phpExcel->createPhpSpreadsheetObject();
        Font::setAutoSizeMethod(Font::AUTOSIZE_METHOD_APPROX);
    }

    /**
     * Initialize the excel writer
     *
     * @param string $filepath
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    public function setup($filepath)
    {
        $this->spreadsheet = $this->phpexcel->createPhpSpreadsheetObject();
        $this->spreadsheet->getActiveSheet()->getPageSetup()->setOrientation(PageSetup::ORIENTATION_LANDSCAPE);
        $this->filepath = $filepath;
    }

    /**
     * Set meta properties for the excel writer
     *
     * @param string $author The author/creator of this file
     * @param string $title  The title of this file
     */
    public function setProperties($author = '', $title = '')
    {
        $this->spreadsheet->getProperties()
            ->setCreator($author)
            ->setTitle($title);
    }

    /**
     * Create a new worksheet, and set it as the active one
     *
     * @param int    $index Location at which to create the sheet (NULL for last)
     * @param string $title The title of the sheet
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    public function setWorksheet($index, $title)
    {
        $this->spreadsheet->createSheet($index);
        $this->resetCurrentRow(1);
        $this->spreadsheet->setActiveSheetIndex($index);
        $this->setSheetTitle($title);
    }

    /**
     * Set the title for the current active worksheet
     *
     * @param string $title
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    public function setSheetTitle($title)
    {
        $this->spreadsheet->getActiveSheet()->setTitle($title);
    }

    /**
     * Write the headers (nested or otherwise) to the current active worksheet
     *
     * @param $headers
     * @param $initRow
     *
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    public function writeHeaders($headers, $initRow = null)
    {
        $worksheet = $this->spreadsheet->getActiveSheet();

        if ($initRow) {
            $worksheet->insertNewRowBefore($initRow, 1);
        } else {
            $initRow = $this->currentRow;
        }

        $column = 'A';
        foreach ($headers as $idx => $header) {
            $cell = $column . $initRow;
            if (!is_array($header)) {
                $worksheet->setCellValue($cell, $header);
                $worksheet->getColumnDimension($column)->setAutoSize(true);
                $column++;
            } else {
                // No more multi-row headers, we are going to flatten nested headers as single header row
                // We only consider "children" for headers ignoring parent header name completely.

                $arrKeys = array_keys($header);
                $headerName = reset($arrKeys);

                // Now write the children's values as flattened header in the row
                foreach ($header[$headerName] as $subHeaderName) {
                    $worksheet->setCellValue($column . $initRow, $subHeaderName);
                    $worksheet->getColumnDimension($column)->setAutoSize(true);
                    $column++;
                }
            }

            // Mark headers as bold
            $worksheet->getStyle($cell)->getFont()->setBold(true);
        }

        $worksheet->calculateColumnWidths();

        $this->currentRow = $initRow + 1;
    }

    /**
     * Check if the given array have multi-row headers
     *
     * @param $headers
     *
     * @return bool True if it has multi-row headers
     */
    private function hasMultiRowHeaders($headers)
    {
        foreach ($headers as $idx => $header) {
            if (is_array($header)) {
                // Multi-row header
                return true;
            }
        }

        return false;
    }

    /**
     * Utility method that will data from the cached file per row and write it
     *
     * @param string $cacheFile Filename where the fetched data can be cached from
     * @param array $sortedHeaders Headers to write sorted in the order you want them
     * @param bool $freezeHeaders True if you want to freeze headers (default: false)
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    public function prepare($cacheFile, $sortedHeaders, $freezeHeaders = false)
    {
        $this->writeHeaders($sortedHeaders);
        if($freezeHeaders) {
            $this->freezePanes();
        }

        $file = new \SplFileObject($cacheFile);
        while (!$file->eof()) {
            $dataRow = json_decode($file->current(), true);

            // Row of data is represented by an array of values. If the $dataRow is a separator string, add a border.
            if (is_string($dataRow) && $dataRow == self::SEPARATOR) {
                $this->addBottomBorder(0, $this->currentRow - 1, count($sortedHeaders) - 1);
            } else {
                $this->writeRow($dataRow, $sortedHeaders);
            }
            $file->next();
        }

        $file = null; // Get rid of the file handle that SplFileObject has on cache file
        unlink($cacheFile);
    }

    public function writeRows($rows, $headers)
    {
        foreach ($rows as $row) {
            $this->writeRow($row, $headers);
        }
    }

    public function writeRow($dataRow, $headers = [])
    {
        if (empty($headers)) {
            // No headers to manage. Just write this array of data directly
            $this->writeArray($dataRow);
            return;
        }

        if (!is_array($dataRow)) {
            // Invalid data -- don't process this row
            return;
        }

        $rowIdx = $this->currentRow;
        $excelRow = [];
        foreach ($headers as $idx => $header) {
            if (!is_array($header)) {
                $excelRow[$rowIdx][] = (isset($dataRow[$header])) ? $dataRow[$header] : null;
            } else {
                // Multi-row header, so we need to set all values
                $nestedHeaderName = array_keys($header)[0];
                $nestedHeaders = $header[$nestedHeaderName];

                foreach ($nestedHeaders as $nestedHeader) {
                    $excelRow[$rowIdx][] = (isset($dataRow[$nestedHeaderName][$nestedHeader])) ? $dataRow[$nestedHeaderName][$nestedHeader] : null;
                }
            }
        }

        $this->writeArrays($excelRow);
        if (isset($dataRow['_bold']) && $dataRow['_bold']) {
            $range = $this->getCellRange(0, $this->currentRow - 1, count($headers) - 1);
            $this->spreadsheet->getActiveSheet()->getStyle($range)->getFont()->setBold(true);
        }
    }

    /**
     * Adds a bottom border to a section defined by numerical coordinates
     *
     * @param $startCol
     * @param $startRow
     * @param $endCol
     * @param $endRow
     */
    private function addBottomBorder($startCol, $startRow, $endCol, $endRow = null)
    {
        $range = $this->getCellRange($startCol, $startRow, $endCol, $endRow);
        $this->spreadsheet->getActiveSheet()->getStyle($range)->getBorders()->getBottom()->setBorderStyle(true);
    }

    /**
     * Converts number coordinates to Excel cell range
     * E.g. getCellRange(0, 5, 9, 6) returns A5:J6
     *
     * @param $startCol
     * @param $startRow
     * @param $endCol
     * @param $endRow
     * @return string
     */
    private function getCellRange($startCol, $startRow, $endCol, $endRow = null)
    {
        return chr(65 + $startCol) . $startRow . ':' . chr(65 + $endCol) . (is_null($endRow) ? $startRow : $endRow);
    }

    /**
     * Write one or more rows starting at the given row and column
     *
     * @param array $lines
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    private function writeArrays(array $lines)
    {
        $startCell = 'A' . $this->currentRow;
        $this->spreadsheet->getActiveSheet()->fromArray($lines, null, $startCell, true);
        $this->currentRow += count($lines);
    }

    /**
     * Write a single row of data
     *
     * @param array $row A single row of data
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    private function writeArray(array $row)
    {
        $startCell = 'A' . $this->currentRow;
        $sheet = $this->spreadsheet->getActiveSheet();
        $sheet->fromArray([$row], null, $startCell, true);

        $column = 'A';
        for ($i = 0; $i < count($row); $i++) {
            if (preg_match("/^=/", $row[$i])) {
                // This is a formula, check it for date & time formulae
                $formats = [];
                if (strpos($row[$i], "DATEVALUE") !== FALSE) {
                    $formats[] = "yyyy-m-d";
                }
                if (strpos($row[$i], "TIMEVALUE") !== FALSE) {
                    $formats[] = "hh:mm:ss";
                }
                if (!empty($formats)) {
                    $sheet->getStyle($column . $this->currentRow)->getNumberFormat()->setFormatCode(implode(" ", $formats));
                }
            }
            $column++;
        }
        $this->currentRow++;
    }

    /**
     * Freeze panes at the given location so they stay fixed upon scroll
     *
     * @param string $cell
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    public function freezePanes($cell = '')
    {
        if (empty($cell)) {
            $cell = 'A2'; // A2 will freeze the rows above cell A2 (i.e row 1)
        }
        $this->spreadsheet->getActiveSheet()->freezePane($cell);
    }

    /**
     * Add a horizonal (row) page break for print layout
     *
     * @param string $cell
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    public function addHorizontalPageBreak($cell = '')
    {
        if (empty($cell)) {
            $cell = 'A' . ($this->currentRow - 2);
        }
        $this->spreadsheet->getActiveSheet()->setBreak($cell, \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet::BREAK_ROW);
    }

    /**
     * Save the current data into an .xlsx file
     *
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    public function finalize()
    {
        // Set active sheet index to the first sheet, so Excel opens this as the first sheet
        $this->spreadsheet->setActiveSheetIndex(0);

        // Write the file to disk
        $writer = $this->phpexcel->createWriter($this->spreadsheet, 'Xlsx');
        $writer->save($this->filepath);
    }

    public function resetCurrentRow($pos)
    {
        $this->currentRow = $pos;
    }

    /**
     * Increment columns. A + 1 = B, A + 2 = C etc...
     *
     * @param     $str
     * @param int $count
     *
     * @return mixed
     */
    private function incrementColumn($str, $count = 1)
    {
        for ($i = 0; $i < $count; $i++) {
            $str++;
        }

        return $str;
    }

    public function getCurrentRow(): int
    {
        return $this->currentRow;
    }
}
