<?php

namespace Tests;

use JackPassero\WindowA3X\WindowA3X;
use PHPUnit\Framework\TestCase;

class WindowWithNotepadA3XTest extends TestCase
{
    private WindowA3X $com;

    private const string TITLE1 = '[REGEXPTITLE:(?i)(window1);]';
    private static string $h1 = '';

    public static function setUpBeforeClass(): void
    {
        $com = new WindowA3X();
        pclose(popen("start /B notepad window1.txt", "r"));
        $com->WinWait(self::TITLE1);
        self::$h1 = $com->WinGetHandle(self::TITLE1);
    }
    public static function tearDownAfterClass(): void
    {
        (new WindowA3X())->WinClose(self::$h1);
    }

    public function setUp(): void
    {
        parent::setUp();
        $this->com = new WindowA3X();
    }

    public function testItCanCreateClass()
    {
        $this->assertNotEmpty(self::$h1);
    }

    public function testWinGetHandle()
    {
        $handle1 = $this->com->WinGetHandle(self::TITLE1);
        $this->assertNotEmpty($handle1);

        $handle2 = $this->com->WinGetHandle(self::TITLE1, 'Окно 1');
        $this->assertNotEmpty($handle2);

        $this->assertSame($handle1, $handle2);

        $handle3 = $this->com->WinGetHandle();
        $this->assertNotEmpty($handle3);
    }

    public function testWin()
    {
        $handle = $this->com->WinGetHandle(self::TITLE1);
        $this->com->WinActivate($handle);
        $this->com->WinWaitActive($handle, 5);
        $this->assertTrue($this->com->WinActive($handle));
        $this->assertTrue($this->com->WinExists($handle));
        $pid = $this->com->WinGetProcess($handle);
        $this->assertSame($pid, $this->com->ProcessExists($pid));
        $text = $this->com->WinGetText($handle);
        $this->assertStringContainsString('Window 1', $text);
        $this->assertStringContainsString('Окно 1', $text);
        $this->assertStringContainsString('window1.txt', $this->com->WinGetTitle($handle));
    }

    public function testWinGetClassList()
    {
        $handle = $this->com->WinGetHandle(self::TITLE1);
        $list = $this->com->WinGetClassList($handle);
        $this->assertNotEmpty($list);
        $this->assertTrue(in_array('NotepadTextBox', $list));
    }

    public function testWinGetState()
    {
        $handle = $this->com->WinGetHandle(self::TITLE1);
        $state = $this->com->WinGetState($handle);
        $this->assertEquals(15, $state);
    }
}
