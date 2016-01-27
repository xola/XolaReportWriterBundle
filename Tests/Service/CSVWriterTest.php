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
        $this->writer->setup("foo.csv");
    }

    public function tearDown()
    {
        $this->writer->finalize();
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

        $contents = file_get_contents("foo.csv");
        $expected = "a,b,\"c,\",d\ne,f,g,h\n";
        $this->assertEquals($expected, $contents);
    }
}
