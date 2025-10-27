<?php

declare(strict_types=1);

//Used as a placeholder in certain COM functions where no parameter is required.
//$empty = new VARIANT();

namespace JackPassero\WindowA3X;

use COM;
use RuntimeException;
use Throwable;
use VARIANT;

class WindowA3X
{
    private COM $com;

    private bool $encode = true;
    private string $encoding = 'CP1251';


    public function __construct()
    {
        if (!extension_loaded('com_dotnet')) {
            throw new RuntimeException('ext-com_dotnet not loaded.');
        }
        try {
            $this->com = new COM("AutoItX3.Control");
        } catch (Throwable) {
            throw new RuntimeException(
                'AutoItX3.Control not registered.'
            );
        }
        /** @noinspection PhpUndefinedMethodInspection */
        $this->com->Opt('WinTitleMatchMode', 4);
    }

    /**
     * @param array<string, mixed> $options
     * @return $this
     */
    public function withOptions(array $options = []): static
    {
        foreach ($options as $key => $val) {
            /** @noinspection PhpUndefinedMethodInspection */
            $this->com->Opt($key, $val);
        }
        return $this;
    }

    public function WinGetHandle(string $title = '[active]', string $text = ''): string
    {
        $ttv = new VARIANT($title, VT_BSTR, CP_UTF8);
        $tev = new VARIANT($text, VT_BSTR, CP_UTF8);
        /** @noinspection PhpUndefinedMethodInspection */
        $handle = $this->com->WinGetHandle($ttv, $tev);
        return hexdec($handle) ? $handle : '';
    }

    public function WinActivate(string $handle): void
    {
        /** @noinspection PhpUndefinedMethodInspection */
        $this->com->WinActivate($this->h($handle));
    }

    public function WinActive(string $handle): bool
    {
        /** @noinspection PhpUndefinedMethodInspection */
        return (bool)$this->com->WinActive($this->h($handle));
    }

    public function WinClose(string $handle): void
    {
        /** @noinspection PhpUndefinedMethodInspection */
        $this->com->WinClose($this->h($handle));
    }

    public function WinExists(string $handle): bool
    {
        /** @noinspection PhpUndefinedMethodInspection */
        return (bool)$this->com->WinExists($this->h($handle));
    }

    public function WinGetPos(string $handle): array
    {
        $h = $this->h($handle);
        /** @noinspection PhpUndefinedMethodInspection */
        return [
            $this->com->WinGetPosX($h),
            $this->com->WinGetPosY($h),
            $this->com->WinGetPosWidth($h),
            $this->com->WinGetPosHeight($h)
        ];
    }

    public function WinGetClientSize(string $handle): array
    {
        $h = $this->h($handle);
        /** @noinspection PhpUndefinedMethodInspection */
        return [
            $this->com->WinGetClientSizeWidth($h),
            $this->com->WinGetClientSizeHeight($h)
        ];
    }

    public function WinGetCaretPos(): array
    {
        /** @noinspection PhpUndefinedMethodInspection */
        return [
            $this->com->WinGetCaretPosX(),
            $this->com->WinGetCaretPosY()
        ];
    }

    /**
     * @param string $handle
     * @return list<string>
     */
    public function WinGetClassList(string $handle): array
    {
        /** @noinspection PhpUndefinedMethodInspection */
        return explode("\n", $this->com->WinGetClassList($this->h($handle)));
    }

    public function WinGetProcess(string $handle): int
    {
        /** @noinspection PhpUndefinedMethodInspection */
        return (int)$this->com->WinGetProcess($this->h($handle));
    }

    public function WinGetState(string $handle): int
    {
        /** @noinspection PhpUndefinedMethodInspection */
        return $this->com->WinGetState($this->h($handle));
    }

    public function WinGetText(string $handle): string
    {
        /** @noinspection PhpUndefinedMethodInspection */
        return $this->toUTF8($this->com->WinGetText($this->h($handle)));
    }

    public function WinGetTitle(string $handle): string
    {
        /** @noinspection PhpUndefinedMethodInspection */
        return $this->toUTF8($this->com->WinGetTitle($this->h($handle)));
    }

    public function WinKill(string $handle): void
    {
        /** @noinspection PhpUndefinedMethodInspection */
        $this->com->WinKill($this->h($handle));
    }


    /**
     * @noinspection PhpUndefinedFieldInspection
     * @noinspection PhpUndefinedMethodInspection
     */
    public function WinList(string $title = '[ALL]', string $text = ''): array
    {
        $vbs = new COM("ScriptControl");
        $vbs->Language = "VBScript";
        $vbs->AddCode('
Function GetWindowsList(title, text)
    Dim oAutoIt,val,result(),i
    Set oAutoIt=CreateObject("AutoItX3.Control")
    val=oAutoIt.WinList(title,text)
    ReDim result(val(0,0)-1)
    For i=1 to val(0,0)
        result(i-1)=Array(val(0,i),val(1,i))
    Next
    GetWindowsList=result
End Function
        ');
        $ttv = new VARIANT($title, VT_BSTR, CP_UTF8);
        $tev = new VARIANT($text, VT_BSTR, CP_UTF8);

        $items = $vbs->Run("GetWindowsList", $ttv, $tev);
        $list = [];
        foreach ($items as $line) {
            $window = [];
            foreach ($line as $i => $item) {
                if ($i === 0) {
                    $item = $this->toUTF8($item);
                }
                $window[] = $item;
            }
            $list[] = $window;
        }

        return $list;
    }


    /**
     * @param string $title
     * @param string $text
     * @return mixed
     * //TODO: Php не может вернуть двумерный массив
     */
    public function WinListNative(string $title = '[ALL]', string $text = ''): mixed
    {
        if (func_num_args() === 0) {
            /** @noinspection PhpUndefinedMethodInspection */
            return $this->com->WinList($title);
        }

        $ttv = new VARIANT($title, VT_BSTR, CP_UTF8);
        $tev = new VARIANT($text, VT_BSTR, CP_UTF8);
        /** @noinspection PhpUndefinedMethodInspection */
        return $this->com->WinList($ttv, $tev);
    }

    public function WinMinimizeAll(): void
    {
        /** @noinspection PhpUndefinedMethodInspection */
        $this->com->WinMinimizeAll();
    }

    public function WinMinimizeAllUndo(): void
    {
        /** @noinspection PhpUndefinedMethodInspection */
        $this->com->WinMinimizeAllUndo();
    }

    /**
     * WinMove has no effect on minimized windows, but WinMove works on hidden windows.
     * If very width and height are small (or negative), the window will go no smaller than 112 x 27 pixels.
     * If width and height are large, the window will go no larger than approximately 12+DesktopWidth x 12+DesktopHeight pixels.
     * Negative values are allowed for the x and y coordinates. In fact, you can move a window off screen; and if the window's program is one that remembers its last window position, the window will appear in the corner (but fully on-screen) the next time you launch the program.
     * @param string $handle
     * @param int $x
     * @param int $y
     * @param int $width
     * @param int $height
     * @return void
     */
    public function WinMove(string $handle, int $x, int $y, int $width = 0, int $height = 0): void
    {
        if ($width === 0 && $height === 0) {
            /** @noinspection PhpUndefinedMethodInspection */
            $this->com->WinMove($this->h($handle), '', $x, $y);
        } else {
            /** @noinspection PhpUndefinedMethodInspection */
            $this->com->WinMove($this->h($handle), '', $x, $y, $width, $height);
        }
    }

    public function WinSetOnTop(string $handle): void
    {
        /** @noinspection PhpUndefinedMethodInspection */
        $this->com->WinSetOnTop($this->h($handle));
    }

    public function WinSetState(string $handle, int $flag): void
    {
        /** @noinspection PhpUndefinedMethodInspection */
        $this->com->WinSetState($this->h($handle), '', $flag);
    }

    public function WinSetTitle(string $handle, string $title): void
    {
        /** @noinspection PhpUndefinedMethodInspection */
        $this->com->WinSetTitle($this->h($handle), '', $title);
    }

    /**
     * Sets the transparency of a window.
     * @param string $handle
     * @param int $transparency A number in the range 0-255. The larger the number, the more transparent the window will become.
     * @return int
     */
    public function WinSetTrans(string $handle, int $transparency): int
    {
        /** @noinspection PhpUndefinedMethodInspection */
        return $this->com->WinSetTrans($this->h($handle), '', $transparency);
    }

    public function WinWaitActive(string $handle, int $timeoutSec = 0): bool
    {
        if ($timeoutSec === 0) {
            /** @noinspection PhpUndefinedMethodInspection */
            return (bool)$this->com->WinWaitActive($this->h($handle));
        }
        /** @noinspection PhpUndefinedMethodInspection */
        return (bool)$this->com->WinWaitActive($this->h($handle), '', $timeoutSec);
    }

    public function WinWaitNotActive(string $handle, int $timeoutSec = 0): bool
    {
        if ($timeoutSec === 0) {
            /** @noinspection PhpUndefinedMethodInspection */
            return (bool)$this->com->WinWaitNotActive($this->h($handle));
        }
        /** @noinspection PhpUndefinedMethodInspection */
        return (bool)$this->com->WinWaitNotActive($this->h($handle), '', $timeoutSec);
    }

    public function WinWait(string $title, string $text = '', int $timeoutSec = 0): bool
    {
        $ttv = new VARIANT($title, VT_BSTR, CP_UTF8);
        $tev = new VARIANT($text, VT_BSTR, CP_UTF8);
        if ($timeoutSec === 0) {
            /** @noinspection PhpUndefinedMethodInspection */
            return (bool)$this->com->WinWait($ttv, $tev);
        }
        /** @noinspection PhpUndefinedMethodInspection */
        return (bool)$this->com->WinWait($ttv, $tev, $timeoutSec);
    }

    public function WinWaitClose(string $handle, int $timeoutSec = 0): bool
    {
        if ($timeoutSec === 0) {
            /** @noinspection PhpUndefinedMethodInspection */
            return (bool)$this->com->WinWaitClose($this->h($handle));
        }
        /** @noinspection PhpUndefinedMethodInspection */
        return (bool)$this->com->WinWaitClose($this->h($handle), '', $timeoutSec);
    }


    public function Send(string $keys, int $flag = 0): void
    {
        /** @noinspection PhpUndefinedMethodInspection */
        $this->com->Send($keys, $flag);
    }

    public function toUTF8(string $text): string|false
    {
        if($this->encode) {
            return iconv($this->encoding, 'UTF-8//IGNORE', $text);
        }
        return $text;
    }

    public function ProcessClose(string|int $process): void
    {
        /** @noinspection PhpUndefinedMethodInspection */
        $this->com->ProcessClose($process);
    }

    public function ProcessExists(string|int $process): int
    {
        /** @noinspection PhpUndefinedMethodInspection */
        return (int)$this->com->ProcessExists($process);
    }

    public function ProcessWait(string|int $process, int $timeoutSec = 0): bool
    {
        /** @noinspection PhpUndefinedMethodInspection */
        return (bool)$this->com->ProcessWait($process, $timeoutSec);
    }

    public function ProcessWaitClose(string|int $process, int $timeoutSec = 0): bool
    {
        /** @noinspection PhpUndefinedMethodInspection */
        return (bool)$this->com->ProcessWaitClose($process, $timeoutSec);
    }

    /**
     * Changes the priority of a process
     *
     * @param string|int $process
     * @param int $priority A flag which determines what priority to set
     * 0 - Idle/Low
     * 1 - Below Normal (Not supported on Windows 95/98/ME)
     * 2 - Normal
     * 3 - Above Normal (Not supported on Windows 95/98/ME)
     * 4 - High
     * 5 - Realtime (Use with caution, may make the system unstable)
     * @return bool
     */
    public function ProcessSetPriority(string|int $process, int $priority): bool
    {
        /** @noinspection PhpUndefinedMethodInspection */
        return (bool)$this->com->ProcessSetPriority($process, $priority);
    }

    public function Run(string $filename, string $workingDir = '', int $flag = 0)
    {
        /** @noinspection PhpUndefinedMethodInspection */
        return $this->com->Run($filename, $workingDir, $flag);
    }

    public function RunWait(string $filename, string $workingDir = '', int $flag = 0)
    {
        /** @noinspection PhpUndefinedMethodInspection */
        return $this->com->RunWait($filename, $workingDir, $flag);
    }

    public function RunAs(
        string $user,
        string $domain,
        string $password,
        int $logonFlag,
        string $filename,
        string $workingDir = '',
        int $flag = 0
    ) {
        /** @noinspection PhpUndefinedMethodInspection */
        return $this->com->RunAs(
            $user,
            $domain,
            $password,
            $logonFlag,
            $filename,
            $workingDir,
            $flag
        );
    }

    public function RunAsWait(
        string $user,
        string $domain,
        string $password,
        int $logonFlag,
        string $filename,
        string $workingDir = '',
        int $flag = 0
    ) {
        /** @noinspection PhpUndefinedMethodInspection */
        return $this->com->RunAsWait(
            $user,
            $domain,
            $password,
            $logonFlag,
            $filename,
            $workingDir,
            $flag
        );
    }

    /**
     * The shutdown code is a combination of the following values:
     * 0 = Logoff
     * 1 = Shutdown
     * 2 = Reboot
     * 4 = Force
     * 8 = Power down
     *
     * Add the required values together. To shutdown and power down, for example, the code would be 9 (shutdown + power down = 1 + 8 = 9).
     *
     * Standby or hibernate can be achieved with third-party software such as http://grc.com/wizmo/wizmo.htm
     * @param int $code
     * @return bool
     */
    public function Shutdown(int $code): bool
    {
        /** @noinspection PhpUndefinedMethodInspection */
        return (bool)$this->com->Shutdown($code);
    }

    public function Sleep(int $milliseconds): void
    {
        /** @noinspection PhpUndefinedMethodInspection */
        $this->com->Sleep($milliseconds);
    }

    public function MouseClick(
        string $button,
        int $x = 0,
        int $y = 0,
        int $clicks = 1,
        int $speed = 10
    ): void {
        /** @noinspection PhpUndefinedMethodInspection */
        $this->com->MouseClick(...func_get_args());
    }

    public function MouseClickDrag(
        string $button,
        int $x1,
        int $y1,
        int $x2,
        int $y2,
        int $speed = 10
    ): void {
        /** @noinspection PhpUndefinedMethodInspection */
        $this->com->MouseClickDrag(...func_get_args());
    }

    /**
     * Returns a cursor ID Number of the current Mouse Cursor.
     * 0 = UNKNOWN (this includes pointing and grabbing hand icons)
     * 1 = APPSTARTING
     * 2 = ARROW
     * 3 = CROSS
     * 4 = HELP
     * 5 = IBEAM
     * 6 = ICON
     * 7 = NO
     * 8 = SIZE
     * 9 = SIZEALL
     * 10 = SIZENESW
     * 11 = SIZENS
     * 12 = SIZENWSE
     * 13 = SIZEWE
     * 14 = UPARROW
     * 15 = WAIT
     * @return int
     */
    public function MouseGetCursor(): int
    {
        /** @noinspection PhpUndefinedMethodInspection */
        return $this->com->MouseGetCursor();
    }

    public function MouseGetPos(): array
    {
        /** @noinspection PhpUndefinedMethodInspection */
        return [
            $this->com->MouseGetPosX(),
            $this->com->MouseGetPosY()
        ];
    }

    public function MouseMove(int $x, int $y, int $speed = 10): int
    {
        /** @noinspection PhpUndefinedMethodInspection */
        return (int)$this->com->MouseMove(...func_get_args());
    }

    public function MouseDown(): void
    {
        /** @noinspection PhpUndefinedMethodInspection */
        $this->com->MouseDown();
    }

    public function MouseUp(): void
    {
        /** @noinspection PhpUndefinedMethodInspection */
        $this->com->MouseUp();
    }

    /**
     * Moves the mouse wheel up or down. XP ONLY.
     * @param string $direction "up" or "down"
     * @param int $clicks Optional: The number of times to move the wheel. Default is 1.
     * @return void
     */
    public function MouseWheel(string $direction, int $clicks = 1): void
    {
        //$ttv = new VARIANT($direction, VT_BSTR, CP_UTF8);
        /** @noinspection PhpUndefinedMethodInspection */
        $this->com->MouseWheel($direction, $clicks);
    }

    public function IsAdmin(): bool
    {
        /** @noinspection PhpUndefinedMethodInspection */
        return (bool)$this->com->IsAdmin();
    }


    public function ClipGet(): mixed
    {
        /** @noinspection PhpUndefinedMethodInspection */
        return $this->com->ClipGet();
    }

    public function ClipPut(string $text): void
    {
        /** @noinspection PhpUndefinedMethodInspection */
        $this->com->ClipPut($text);
    }

    public function PixelChecksum(
        int $left,
        int $top,
        int $right,
        int $bottom,
        int $step = 1
    ): float {
        /** @noinspection PhpUndefinedMethodInspection */
        return (float)$this->com->PixelChecksum(...func_get_args());
    }

    public function PixelGetColor(int $x, int $y): int
    {
        /** @noinspection PhpUndefinedMethodInspection */
        return $this->com->PixelGetColor($x, $y);
    }

    public function PixelSearch(
        int $left,
        int $top,
        int $right,
        int $bottom,
        int $color,
        int $shadeVariation = 0,
        int $step = 1,
    ): array {
        /** @noinspection PhpUndefinedMethodInspection */
        $variant = $this->com->PixelSearch(...func_get_args());
        if ($this->GetAutoItError()) {
            return [];
        }

        $position = [];
        foreach ($variant as $item) {
            $position[] = $item;
        }
        return $position;
    }

    public function GetAutoItError(): mixed
    {
        /** @noinspection PhpUndefinedMethodInspection */
        return $this->com->error();
    }

    public function ControlClick(
        string $handle,
        string $controlID,
        string $button = 'left',
        int $clicks = 1,
        $x = 0,
        $y = 0
    ): bool {
        if ($x === 0 && $y === 0) {
            /** @noinspection PhpUndefinedMethodInspection */
            return (bool)$this->com->ControlClick(
                $this->h($handle),
                '',
                $controlID,
                $button,
                $clicks
            );
        }

        /** @noinspection PhpUndefinedMethodInspection */
        return (bool)$this->com->ControlClick(
            $this->h($handle),
            '',
            $controlID,
            $button,
            $clicks,
            $x,
            $y
        );
    }

    public function ControlCommand(
        string $handle,
        string $controlID,
        string $command,
        string $option = '',
    ): void {
        /** @noinspection PhpUndefinedMethodInspection */
        $this->com->ControlCommand($this->h($handle), '', $controlID, $command, $option);
    }

    public function ControlFocus(string $handle, string $controlID = ''): int
    {
        /** @noinspection PhpUndefinedMethodInspection */
        return $this->com->ControlFocus($this->h($handle), '', $controlID);
    }

    public function ControlSend(
        string $handle,
        string $string,
        string $controlID = '',
        int $flag = 0
    ): int {
        /** @noinspection PhpUndefinedMethodInspection */
        return $this->com->ControlSend($this->h($handle), '', $controlID, $string, $flag);
    }

    public function setEncoding(string $encoding): void
    {
        $this->encoding = $encoding;
    }

    public function setEncode(bool $encode): void
    {
        $this->encode = $encode;
    }


    private function h(string $handle): string
    {
        return '[HANDLE:' . $handle . ']';
    }
}
