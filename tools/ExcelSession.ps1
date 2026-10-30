# Excel を自動操作で起動・終了するための道具。build.ps1 と tests\run-tests.ps1 から読み込みます。
# ・起動した Excel のプロセス番号を覚えておき、終わるときはその Excel だけを止めます
#   (ほかに開いている Excel には触りません)。
# ・VBA のコンパイルエラーなどで Excel が固まったときに備えて、時間の見張りをつけます。

if (-not ('BatchInvoicePdf.Build.Win32' -as [type])) {
    Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
namespace BatchInvoicePdf.Build {
    public static class Win32 {
        [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);
    }
    // PowerShell 5.1 から COM を呼ぶと、同じ行で型の違う値(文字のあとに数値など)を渡したときに
    // 型の変換で失敗することがあるため、値を渡す呼び出しはここを通す。
    public static class Com {
        static object Unwrap(object value) {
            var wrapped = value as System.Management.Automation.PSObject;
            return wrapped == null ? value : wrapped.BaseObject;
        }

        // Range.Value2 に値(1つの値、または2次元配列)を入れる
        public static void SetValue2(object range, object value) {
            range.GetType().InvokeMember("Value2", System.Reflection.BindingFlags.SetProperty, null, range, new object[] { Unwrap(value) });
        }

        // Application.Run でマクロを呼び、戻り値を返す
        public static object Run(object app, string macro, object[] args) {
            var all = new object[args.Length + 1];
            all[0] = macro;
            for (int i = 0; i < args.Length; i++) all[i + 1] = Unwrap(args[i]);
            return app.GetType().InvokeMember("Run", System.Reflection.BindingFlags.InvokeMethod, null, app, all);
        }
    }
}
'@
}

function Start-ExcelSession {
    param([int]$TimeoutSec = 900)

    $app = New-Object -ComObject Excel.Application
    $processId = [uint32]0
    [void][BatchInvoicePdf.Build.Win32]::GetWindowThreadProcessId([IntPtr]$app.Hwnd, [ref]$processId)
    $app.Visible = $false
    $app.DisplayAlerts = $false
    $app.ScreenUpdating = $false
    # プロセス番号は使い回されるので、起動時刻も覚えておき、止める前に同じ Excel かを確かめる
    $startTicks = (Get-Process -Id $processId).StartTime.Ticks

    # 時間内に終わらなければ、この Excel だけを止める
    $watchdog = Start-Job -ScriptBlock {
        param($targetId, $targetTicks, $seconds)
        Start-Sleep -Seconds $seconds
        $p = Get-Process -Id $targetId -ErrorAction SilentlyContinue
        if ($p -and $p.ProcessName -eq 'EXCEL' -and $p.StartTime.Ticks -eq $targetTicks) {
            Stop-Process -Id $targetId -Force -ErrorAction SilentlyContinue
        }
    } -ArgumentList $processId, $startTicks, $TimeoutSec

    [pscustomobject]@{ App = $app; ProcessId = [int]$processId; StartTicks = $startTicks; Watchdog = $watchdog }
}

# 起動した Excel がまだ動いているか(同じ番号の別のプロセスは数えない)
function Test-ExcelSessionAlive {
    param($Session)
    $p = Get-Process -Id $Session.ProcessId -ErrorAction SilentlyContinue
    [bool]($p -and $p.ProcessName -eq 'EXCEL' -and $p.StartTime.Ticks -eq $Session.StartTicks)
}

function Stop-ExcelSession {
    param($Session)

    if ($null -eq $Session) { return }
    try {
        foreach ($wb in @($Session.App.Workbooks)) { $wb.Close($false) }
        $Session.App.Quit()
    } catch {
        # 見張りに止められたあとなどは、ここで失敗してもかまわない
    }
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($Session.App)
    Stop-Job -Job $Session.Watchdog -ErrorAction SilentlyContinue
    Remove-Job -Job $Session.Watchdog -Force -ErrorAction SilentlyContinue

    # Quit のあとも少し残ることがあるので、待ってから、この Excel だけを止める
    $deadline = (Get-Date).AddSeconds(10)
    while ((Test-ExcelSessionAlive $Session) -and ((Get-Date) -lt $deadline)) {
        Start-Sleep -Milliseconds 300
    }
    if (Test-ExcelSessionAlive $Session) {
        Stop-Process -Id $Session.ProcessId -Force -ErrorAction SilentlyContinue
    }
}

# Excel の色の値(赤 + 緑*256 + 青*65536)
function ConvertTo-ExcelColor {
    param([int]$Red, [int]$Green, [int]$Blue)
    $Red + 256 * $Green + 65536 * $Blue
}
