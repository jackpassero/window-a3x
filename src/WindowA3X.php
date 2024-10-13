<?php

declare(strict_types=1);

namespace JackPassero\WindowA3X;

use COM;
use RuntimeException;
use Throwable;
use VARIANT;

class WindowA3X
{
    private COM $com;

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

        $this->com->Opt('WinTitleMatchMode', 4);
    }

    /**
     * @param array<string, mixed> $options
     * @return $this
     */
    public function withOptions(array $options = []): static
    {
        foreach ($options as $key => $val) {
            $this->com->Opt($key, $val);
        }
        return $this;
    }

    public function send(string $keys, int $flag = 0): void
    {
        $this->com->Send($keys, $flag);
    }

    public function WinGetHandle(string $title = null): string
    {
        $v = new VARIANT($title, VT_BSTR, CP_UTF8);
        $h = $this->com->WinGetHandle($v);
        return hexdec($h) ? $h : '';
    }

    public function WinWait(string $title, int $timeout = 0): bool
    {
        $v = new VARIANT($title, VT_BSTR, CP_UTF8);
        return (bool)$this->com->WinWait($v, '', $timeout);
    }

    public function WinWaitClose(string $title, int $timeout = 0): bool
    {
        $v = new VARIANT($title, VT_BSTR, CP_UTF8);
        return (bool)$this->com->WinWaitClose($v, '', $timeout);
    }

    public function WinWaitCloseByHandle(string $handle, int $timeout = 0): bool
    {
        return (bool)$this->com->WinWaitClose($this->handle($handle), '', $timeout);
    }

    public function WinExists(string $handle): bool
    {
        return (bool)$this->com->WinExists($this->handle($handle));
    }

    public function WinGetTitle(string $handle): string
    {
        return $this->toUTF8($this->com->WinGetTitle($this->handle($handle)));
    }

    public function WinCloseByHandle(string $handle): bool
    {
        return (bool)$this->com->WinClose($this->handle($handle));
    }

    public function ControlSend(string $handle, string $string, int $flag = 0): int
    {
        return $this->com->ControlSend($this->handle($handle), '', '', $string, $flag);
    }

    public function ControlFocus(string $handle): int
    {
        return $this->com->ControlFocus($this->handle($handle), '', '');
    }

    public function toUTF8(string $title): string|false
    {
        return iconv('WINDOWS-1251', 'UTF-8//IGNORE', $title);
    }

    public function ProcessClose(string|int $process): int
    {
        return $this->com->ProcessClose($process);
    }

    public function ProcessExists(string|int $process): int
    {
        return $this->com->ProcessExists($process);
    }

    public function ProcessWaitClose(string|int $process, int $timeout = 0): int
    {
        return $this->com->ProcessWaitClose($process, $timeout);
    }

    public function WinKillByHandle(string $handle): bool
    {
        return (bool)$this->com->WinKill($this->handle($handle));
    }

    private function handle(string $handle): string
    {
        return '[HANDLE:' . $handle . ']';
    }
}
