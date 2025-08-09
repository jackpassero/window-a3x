<?php

namespace Tests;

use JackPassero\WindowA3X\WindowA3X;
use PHPUnit\Framework\TestCase;

class WindowA3XTest extends TestCase
{
    private WindowA3X $com;


    public function setUp(): void
    {
        parent::setUp();
        $this->com = new WindowA3X();
    }

    public function test_it_can_create_class()
    {
        $this->assertInstanceOf(WindowA3X::class, $this->com);
    }

    public function test_it_can_process_exists()
    {
        $pid = getmypid();
        $this->assertSame($pid, $this->com->ProcessExists($pid));
    }

}
