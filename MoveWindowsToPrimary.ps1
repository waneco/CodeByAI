# PowerShellスクリプト: プライマリモニター以外のウィンドウをプライマリモニターへ移動

# 必要なWinAPIを定義
Add-Type @"
using System;
using System.Runtime.InteropServices;

public class WinAPI
{
    [DllImport("user32.dll")]
    public static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

    public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

    [DllImport("user32.dll")]
    public static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

    [DllImport("user32.dll")]
    public static extern bool IsWindowVisible(IntPtr hWnd);

    [DllImport("user32.dll")]
    public static extern bool MoveWindow(IntPtr hWnd, int X, int Y, int nWidth, int nHeight, bool bRepaint);

    [StructLayout(LayoutKind.Sequential)]
    public struct RECT
    {
        public int Left;
        public int Top;
        public int Right;
        public int Bottom;
    }
}
"@

# プライマリモニターの解像度と位置を取得
Add-Type -AssemblyName System.Windows.Forms
$PrimaryScreen = [System.Windows.Forms.Screen]::PrimaryScreen
$PrimaryMonitorLeft = $PrimaryScreen.Bounds.X
$PrimaryMonitorTop = $PrimaryScreen.Bounds.Y
$PrimaryMonitorWidth = $PrimaryScreen.Bounds.Width
$PrimaryMonitorHeight = $PrimaryScreen.Bounds.Height

Write-Host "プライマリモニターの座標: 左=$PrimaryMonitorLeft, 上=$PrimaryMonitorTop, 幅=$PrimaryMonitorWidth, 高さ=$PrimaryMonitorHeight"

# ウィンドウのタイトルを取得する関数
function Get-WindowTitle {
    param ($hWnd)
    Add-Type @"
    using System;
    using System.Runtime.InteropServices;
    public class WindowTitleAPI {
        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        public static extern int GetWindowText(IntPtr hWnd, System.Text.StringBuilder lpString, int nMaxCount);
    }
"@
    $sb = New-Object System.Text.StringBuilder 256
    [WindowTitleAPI]::GetWindowText($hWnd, $sb, $sb.Capacity) | Out-Null
    return $sb.ToString()
}

# ウィンドウをプライマリモニターに移動する関数
function MoveWindowsToPrimaryMonitor {
    param ($hWnd, $PrimaryMonitorLeft, $PrimaryMonitorTop, $PrimaryMonitorWidth, $PrimaryMonitorHeight)

    # ウィンドウが表示中か確認
    if (-not [WinAPI]::IsWindowVisible($hWnd)) {
        Write-Verbose "非表示ウィンドウをスキップ: ハンドル=$hWnd"
        return
    }

    # ウィンドウタイトルを取得
    $windowTitle = Get-WindowTitle $hWnd
    if ([string]::IsNullOrWhiteSpace($windowTitle)) {
        $windowTitle = "無題"
    }

    # 特定のウィンドウを無視する例
    if ($windowTitle -match "Taskbar" -or $windowTitle -match "System Tray") {
        Write-Verbose "無視されたウィンドウ: タイトル='$windowTitle', ハンドル=$hWnd"
        return
    }

    # ウィンドウの位置を取得
    $rect = New-Object WinAPI+RECT
    try {
        if (-not [WinAPI]::GetWindowRect($hWnd, [ref]$rect)) {
            throw "座標取得に失敗しました: ハンドル=$hWnd"
        }
    } catch {
        Write-Error $_
        return
    }

    $width = $rect.Right - $rect.Left
    $height = $rect.Bottom - $rect.Top

    # ウィンドウがプライマリモニター外にあるかを判定
    $isOutsidePrimaryMonitor = $rect.Left -lt $PrimaryMonitorLeft -or $rect.Right -gt ($PrimaryMonitorLeft + $PrimaryMonitorWidth) -or `
        $rect.Top -lt $PrimaryMonitorTop -or $rect.Bottom -gt ($PrimaryMonitorTop + $PrimaryMonitorHeight)

    if ($isOutsidePrimaryMonitor) {
        # プライマリモニターに移動
        $newLeft = $PrimaryMonitorLeft + [Math]::Max(0, $rect.Left - $PrimaryMonitorLeft)
        $newTop = $PrimaryMonitorTop + [Math]::Max(0, $rect.Top - $PrimaryMonitorTop)
        $newWidth = [Math]::Min($width, $PrimaryMonitorWidth)
        $newHeight = [Math]::Min($height, $PrimaryMonitorHeight)

        [WinAPI]::MoveWindow($hWnd, $newLeft, $newTop, $newWidth, $newHeight, $true)
        Write-Host "ウィンドウを移動しました: タイトル='$windowTitle', ハンドル=$hWnd"
    } else {
        Write-Verbose "プライマリモニター上のウィンドウ: タイトル='$windowTitle', ハンドル=$hWnd"
    }
}

# ウィンドウを列挙して処理
$movedCount = 0
[WinAPI]::EnumWindows({
    param ($hWnd, $lParam)
    MoveWindowsToPrimaryMonitor $hWnd $PrimaryMonitorLeft $PrimaryMonitorTop $PrimaryMonitorWidth $PrimaryMonitorHeight
    $movedCount++
    return $true
}, [IntPtr]::Zero)

Write-Host "処理が完了しました。移動したウィンドウ数: $movedCount"
