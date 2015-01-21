<?php

namespace Xola\ReportWriterBundle\Service;

use Symfony\Component\DependencyInjection\Container;
use Psr\Log\LoggerInterface;

class CSVWriter
{
    protected static $headers = [];
    protected $logger;
    private $cacheFile;

    public function __construct(Container $container, LoggerInterface $logger)
    {
        $this->logger = $logger;
        $this->cacheFile = tempnam($container->get('kernel')->getCacheDir(), 'data_export');
    }

    /**
     * Write the formatted order data to disk, so we can fetch it later
     *
     * @param $data
     */
    public function cacheData($data)
    {
        $line = json_encode($data) . "\n";
        file_put_contents($this->cacheFile, $line, FILE_APPEND);
    }

    /**
     * Go through the order data and prepare an updated list of headers
     *
     * @param $data
     *
     * @return array
     */
    public function parseHeaders($data)
    {
        foreach ($data as $key => $value) {
            if (is_array($value)) {
                // This is a multi-row header
                $value = array_keys($value);
                $loc = $this->findNestedHeader($key);
                if ($loc !== false) {
                    // Merge data headers values with pre-existing data
                    $value = array_unique(array_merge($value, self::$headers[$loc][$key]));
                    self::$headers[$loc] = [$key => $value];
                } else {
                    self::$headers[] = [$key => $value];
                }

            } else {
                // Standard header add it if it does not exist
                if (!in_array($key, self::$headers)) {
                    self::$headers[] = $key;
                }
            }
        }

        return self::$headers;
    }

    /**
     * Find the location of a nested/multi-row header from our list
     *
     * @param string $key
     *
     * @return bool|int FALSE if it is not found, else location of the header within the array
     */
    private function findNestedHeader($key)
    {
        $found = false;
        foreach (self::$headers as $idx => $value) {
            if (is_array($value) && isset($value[$key])) {
                return $idx;
            }
        }

        return $found;
    }

    /**
     * Write the compiled csv to disk and return the file name
     *
     * @param array $sortedHeaders An array of sorted headers
     *
     * @return string The csv filename where the data was written
     */
    public function prepare($sortedHeaders)
    {
        $csvFile = $this->cacheFile . '.csv';
        $handle = fopen($csvFile, 'w');

        // Generate a csv version of the multi-row headers to write to disk
        $headerRows = [];
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
                        $headerRows[0] = array_merge($headerRows[0], array_fill(0, count($subHeaders)-1, ''));
                    }
                }
            }
        }

        fputcsv($handle, $headerRows[0]);
        fputcsv($handle, $headerRows[1]);

        // TODO: Track memory usage
        $file = new \SplFileObject($this->cacheFile);
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
        unlink($this->cacheFile);

        fclose($handle);

        return $csvFile;
    }
}