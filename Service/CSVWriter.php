<?php

namespace Xola\ReportWriterBundle\Service;

class CSVWriter extends AbstractWriter
{
    /**
     * Write the compiled csv to disk and return the file name
     *
     * @param array $sortedHeaders An array of sorted headers
     *
     * @return string The csv filename where the data was written
     */
    public function prepare($cacheFile, $sortedHeaders)
    {
        $csvFile = $cacheFile . '.csv';
        $handle = fopen($csvFile, 'w');

        // Generate a csv version of the multi-row headers to write to disk
        $headerRows = [[], []];
        foreach ($sortedHeaders as $idx => $header) {
            if (!is_array($header)) {
                $headerRows[0][] = $header;
                $headerRows[1][] = '';
            } else {
                foreach ($header as $headerName => $subHeaders) {
                    $headerRows[0][] = $headerName;
                    $headerRows[1] = array_merge($headerRows[1], $subHeaders);
                    if (count($subHeaders) > 1) {
                        /**
                         * We need to insert empty cells for the first row of headers to account for the second row
                         * this acts as a faux horizontal cell merge in a csv file
                         * | Header 1 | <---- 2 extra cells ----> |
                         * | Sub 1    | Subheader 2 | Subheader 3 |
                         */
                        $headerRows[0] = array_merge($headerRows[0], array_fill(0, count($subHeaders) - 1, ''));
                    }
                }
            }
        }

        fputcsv($handle, $headerRows[0]);
        fputcsv($handle, $headerRows[1]);

        // TODO: Track memory usage
        $file = new \SplFileObject($cacheFile);
        while (!$file->eof()) {
            $csvRow = [];
            $row = json_decode($file->current(), true);
            if (!is_array($row)) {
                // Invalid json data -- don't process this row
                continue;
            }
            foreach ($sortedHeaders as $idx => $header) {
                if (!is_array($header)) {
                    $csvRow[] = (isset($row[$header])) ? $row[$header] : '';
                } else {
                    // Multi-row header, so we need to set all values
                    $nestedHeaderName = array_keys($header)[0];
                    $nestedHeaders = $header[$nestedHeaderName];

                    foreach ($nestedHeaders as $nestedHeader) {
                        $csvRow[] = (isset($row[$nestedHeaderName][$nestedHeader])) ? $row[$nestedHeaderName][$nestedHeader] : '';
                    }
                }
            }

            fputcsv($handle, $csvRow);
            $file->next();
        }

        $file = null; // Get rid of the file handle that SplFileObject has on cache file
        unlink($cacheFile);

        fclose($handle);

        return $csvFile;
    }
}