<?php

namespace Tests;

use JackPassero\WindowA3X\WindowA3X;
use PHPUnit\Framework\TestCase;

class OtherA3XTest extends TestCase
{
    private WindowA3X $com;

    public function setUp(): void
    {
        parent::setUp();
        $this->com = new WindowA3X();
    }

    public function testRun()
    {
        $pid = $this->com->Run('calc.exe');
        $this->assertSame($pid, $this->com->ProcessExists($pid));
    }

    public function testMouseClick()
    {
        $this->com->MouseClick('left', clicks: 2);
    }

    public function testMouseGetCursor()
    {
        $id = $this->com->MouseGetCursor();
        $this->assertGreaterThanOrEqual(0, $id);
    }

    public function testMouse()
    {
        $this->com->MouseMove(100, 100, 0);
        $pos = $this->com->MouseGetPos();
        $this->assertGreaterThanOrEqual(100, $pos[0]);
        $this->assertGreaterThanOrEqual(100, $pos[1]);

        $this->com->MouseDown();
        $this->com->MouseUp();

    }

    public function testMouseWheel()
    {
        $this->com->MouseWheel('up', 2);
        $this->assertSame('up', 'up');
    }

    public function testMisc()
    {
        $this->assertFalse($this->com->IsAdmin());

        $this->com->ClipPut($clip = 'Привет');
        $text = $this->com->ClipGet();
        $this->assertSame($clip, $text);

    }

    public function testPixel()
    {
        $a = $this->com->PixelChecksum(0, 0, 0, 0);
        $this->assertGreaterThan(0, $a);

        $color = $this->com->PixelGetColor(5, 5);
        $data = $this->com->PixelSearch(0, 0, 10, 10, $color);
        $this->assertCount(2, $data);
        $this->assertEquals(0, $this->com->GetAutoItError());
        $this->com->PixelSearch(0, 0, 10, 10, 0x831233);
        $this->assertEquals(1, $this->com->GetAutoItError());
    }

    public function testControlGetHandle()
    {
        $hw = $this->com->WinGetHandle();
        $hc = $this->com->ControlGetHandle($hw);
        $this->assertNotEmpty($hc);
    }


}
