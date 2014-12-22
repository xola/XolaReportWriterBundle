<?php

namespace Xola\ReportWriterBundle\Service;

use Symfony\Component\DependencyInjection\Container;
use Psr\Log\LoggerInterface;

class CSVWriter
{
    protected static $headers = [];
    protected $logger;

    public function __construct(Container $container, LoggerInterface $logger)
    {
        $this->logger = $logger;
    }

    public function cacheOrders($data)
    {
        return $data;
    }
}