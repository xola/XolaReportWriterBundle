<?php

use Xola\ReportWriterBundle\Service\CSVWriter;

class CSVWriterTest extends PHPUnit_Framework_TestCase
{
    /** @var CSVWriter */
    private $writer;
    private $handle;

    public function setUp()
    {
        $this->writer = $this->buildService();
        $this->handle = fopen('php://temp', 'r+');
        $this->writer->setHandle($this->handle);
    }

    public function tearDown()
    {
        fclose($this->handle);
        $this->handle = null;
    }

    public function buildService($params = [])
    {
        $defaults = ['logger' => $this->getMockBuilder('Psr\Log\LoggerInterface')->disableOriginalConstructor()->getMock()];
        $params = array_merge($defaults, $params);

        return new CSVWriter($params['logger']);
    }

    public function testShouldWriteRowToFile()
    {
        $this->writer->writeRow(['a', 'b', 'c,', 'd']);
        $this->writer->writeRow(['e', 'f', 'g', 'h']);

        $contents = $this->getFileContents();
        $expected = "a,b,\"c,\",d\ne,f,g,h\n";
        $this->assertEquals($expected, $contents);
    }

    private function getFileContents()
    {
        rewind($this->handle);
        return stream_get_contents($this->handle);
    }
}
