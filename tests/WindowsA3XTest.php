<?php

namespace Tests;

use JackPassero\WindowA3X\WindowA3X;
use PHPUnit\Framework\TestCase;

class WindowsA3XTest extends TestCase
{
    private WindowA3X $com;

    public function setUp(): void
    {
        parent::setUp();
        $this->com = new WindowA3X();
    }

    public function testWinList()
    {
        $list = $this->com->WinList();
        $this->assertGreaterThan(0, count($list));
    }

    public function testWinListRus()
    {
        $this->com->setEncode(true);
        $this->com->setEncoding('Windows-1251');
        $list = $this->com->WinList('[REGEXPTITLE:(?i)(Русский)]');
        $this->assertGreaterThan(0, count($list));
    }


}
