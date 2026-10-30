<#
.SYNOPSIS
    できた .xlsm を Excel で動かして、決まりどおりに動くかを確かめます。

.DESCRIPTION
    リポジトリの .xlsm はそのままにして、一時フォルダ(%TEMP%)にコピーしたものを使います。
    マクロは画面を出さない本体(modMain.ExecuteBatch / ExecutePreview)を Application.Run で呼びます。
    確かめること:
      ・エラー確認用のデータで、決めたとおりのエラーが「エラー一覧」に出る(tests\expected-errors.csv)
      ・正常なデータで、エラーがない/20件のPDFができる/ファイル名が決まりどおり/1ページずつ
      ・ひな形の小計・消費税・合計が、CSV から計算した値と合う(端数処理の3つの設定すべて)
      ・エラーが1件でもあれば、1件も作らない
      ・同名ファイル:「止める」で止まり、「上書き」で置き換わる
      ・想定外のエラー(PDFを開いたまま上書き)で、止まった請求番号が画面の文と実行履歴に出る
      ・実行履歴が残り、ブックを閉じて開き直しても消えない
      ・ボタンにマクロがついている。配る状態のブックに、作成者名や実行の跡が残っていない

.EXAMPLE
    powershell -NoProfile -ExecutionPolicy Bypass -File tests\run-tests.ps1
#>
param(
    [string]$BookDir = (Split-Path -Parent $PSScriptRoot),
    [switch]$KeepWork
)

$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
. (Join-Path $root 'tools\ExcelSession.ps1')
Add-Type -AssemblyName System.IO.Compression, System.IO.Compression.FileSystem

$normalBook = Join-Path $BookDir '請求書一括PDF作成.xlsm'
$errorBook = Join-Path $BookDir '請求書一括PDF作成_エラー確認用.xlsm'
$work = Join-Path ([IO.Path]::GetTempPath()) ('batch-invoice-pdf-test-' + [Guid]::NewGuid().ToString('N').Substring(0, 8))
New-Item -ItemType Directory -Path $work | Out-Null

$results = New-Object 'System.Collections.Generic.List[object]'
function Test-That {
    param([string]$Name, [bool]$Ok, [string]$Detail = '')
    $results.Add([pscustomobject]@{ Name = $Name; Ok = $Ok; Detail = $Detail })
    if ($Ok) { Write-Host "OK  $Name" } else { Write-Host "NG  $Name … $Detail" }
}

#-------------------------------------------------------------------------------
# 道具
#-------------------------------------------------------------------------------
function Open-Copy {
    param([string]$Source, [string]$Name)
    $dest = Join-Path $work $Name
    Copy-Item -LiteralPath $Source -Destination $dest -Force
    $app.Workbooks.Open($dest)
}

function Invoke-Macro {
    param($Workbook, [string]$Macro, [object[]]$Arguments = @())
    [BatchInvoicePdf.Build.Com]::Run($app, "'" + $Workbook.Name + "'!" + $Macro, [object[]]$Arguments)
}

function Set-Setting {
    param($Workbook, [string]$Label, $Value)
    $ws = $Workbook.Worksheets.Item('設定')
    for ($r = 1; $r -le 60; $r++) {
        if ($ws.Cells.Item($r, 1).Text -eq $Label) {
            [BatchInvoicePdf.Build.Com]::SetValue2($ws.Cells.Item($r, 2), $Value)
            return
        }
    }
    throw "設定シートに「$Label」がありません"
}

function Get-ErrorRows {
    param($Workbook)
    $ws = $Workbook.Worksheets.Item('エラー一覧')
    $rows = @()
    for ($r = 5; $r -le 500; $r++) {
        if ($ws.Cells.Item($r, 1).Text -eq '') { break }
        $rows += [pscustomobject]@{
            Key     = ($ws.Cells.Item($r, 2).Text, $ws.Cells.Item($r, 3).Text, $ws.Cells.Item($r, 4).Text, $ws.Cells.Item($r, 5).Text) -join ','
            Message = $ws.Cells.Item($r, 6).Text
        }
    }
    , $rows
}

function Get-History {
    param($Workbook)
    $ws = $Workbook.Worksheets.Item('実行履歴')
    $rows = @()
    for ($r = 4; $r -le 500; $r++) {
        if ($ws.Cells.Item($r, 1).Text -eq '') { break }
        $rows += [pscustomobject]@{
            Operation = $ws.Cells.Item($r, 2).Text; Result = $ws.Cells.Item($r, 3).Text
            Count = $ws.Cells.Item($r, 4).Text; Folder = $ws.Cells.Item($r, 5).Text; Detail = $ws.Cells.Item($r, 6).Text
        }
    }
    , $rows
}

function Get-PdfInfo {
    param([string]$Path)
    $text = [Text.Encoding]::GetEncoding(28591).GetString([IO.File]::ReadAllBytes($Path))
    [pscustomobject]@{
        IsPdf  = $text.StartsWith('%PDF-')
        Pages  = ([regex]::Matches($text, '/Type\s*/Page(?!s)')).Count
        Author = [regex]::IsMatch($text, '/Author')
    }
}

# CSV から、請求ごとの小計・消費税・合計を計算する(ひな形の式と同じ決まり)
function Get-ExpectedTotals {
    param([string]$Rounding)
    $round = {
        param([decimal]$x)
        switch ($Rounding) {
            '切り捨て' { [decimal][math]::Truncate($x) }
            '切り上げ' { if ($x -ge 0) { [decimal][math]::Ceiling($x) } else { [decimal][math]::Floor($x) } }
            '四捨五入' { [decimal][math]::Round($x, 0, [MidpointRounding]::AwayFromZero) }
        }
    }
    $totals = @{}
    foreach ($d in Import-Csv -Path (Join-Path $root 'testdata\正常\明細.csv') -Encoding UTF8) {
        if (-not $totals.ContainsKey($d.'請求番号')) { $totals[$d.'請求番号'] = @{ S10 = [decimal]0; S8 = [decimal]0 } }
        $amount = [decimal][math]::Truncate([decimal]$d.'数量' * [decimal]$d.'単価')
        if ($d.'税率' -eq '10%') { $totals[$d.'請求番号'].S10 += $amount } else { $totals[$d.'請求番号'].S8 += $amount }
    }
    foreach ($k in @($totals.Keys)) {
        $t = $totals[$k]
        $t.T10 = & $round ($t.S10 * 10 / 100)
        $t.T8 = & $round ($t.S8 * 8 / 100)
        $t.Total = $t.S10 + $t.T10 + $t.S8 + $t.T8
    }
    $totals
}

#-------------------------------------------------------------------------------
# 配る状態のブック(リポジトリのファイルそのものを読むだけ)
#-------------------------------------------------------------------------------
foreach ($book in @($normalBook, $errorBook)) {
    $zip = [IO.Compression.ZipFile]::OpenRead($book)
    try {
        $core = New-Object System.IO.StreamReader($zip.GetEntry('docProps/core.xml').Open())
        $xml = [xml]$core.ReadToEnd()
        $core.Close()
        $creator = @($xml.GetElementsByTagName('dc:creator') | ForEach-Object { $_.InnerText }) -join ''
        $lastBy = @($xml.GetElementsByTagName('cp:lastModifiedBy') | ForEach-Object { $_.InnerText }) -join ''
        Test-That "配る状態:作成者名が残っていない($(Split-Path -Leaf $book))" ($creator -eq '' -and $lastBy -eq '') "creator=$creator lastModifiedBy=$lastBy"
        Test-That "配る状態:VBA が入っている($(Split-Path -Leaf $book))" ($null -ne $zip.GetEntry('xl/vbaProject.bin'))
    } finally { $zip.Dispose() }
}

$session = Start-ExcelSession -TimeoutSec 900
$app = $session.App
try {
    #---------------------------------------------------------------------------
    # 配る状態:ひな形が空、エラー一覧・実行履歴がまっさら、ボタンにマクロ
    #---------------------------------------------------------------------------
    $wb = $app.Workbooks.Open($normalBook, 0, $true)   # 読み取り専用で開く
    $tpl = $wb.Worksheets.Item('請求書ひな形')
    Test-That '配る状態:ひな形の差し込み先が空' ($tpl.Range('F4').Text -eq '' -and $tpl.Range('B14').Text -eq '' -and $tpl.Range('B4').Text -eq '')
    Test-That '配る状態:エラー一覧が未チェック' ($wb.Worksheets.Item('エラー一覧').Range('A2').Text -eq 'まだチェックしていません。')
    Test-That '配る状態:実行履歴が空' ((Get-History $wb).Count -eq 0)
    $actions = @($wb.Worksheets.Item('使い方').Shapes | ForEach-Object { $_.OnAction })
    foreach ($m in 'modMain.RunCheckOnly', 'modMain.RunCreatePdfs', 'modMain.RunPreview') {
        Test-That "ボタン:$m がついている" (@($actions | Where-Object { $_ -like "*$m" }).Count -eq 1) ($actions -join ' / ')
    }
    Test-That '配る状態:最初に「使い方」シートが開く' ($wb.ActiveSheet.Name -eq '使い方') $wb.ActiveSheet.Name
    $wb.Close($false)

    #---------------------------------------------------------------------------
    # エラー確認用のデータ
    #---------------------------------------------------------------------------
    $wb = Open-Copy $errorBook 'errors.xlsm'
    $message = Invoke-Macro $wb 'modMain.ExecuteBatch' @($false)
    $expected = @(Import-Csv -Path (Join-Path $PSScriptRoot 'expected-errors.csv') -Encoding UTF8 |
        ForEach-Object { ($_.'シート', $_.'行', $_.'請求番号', $_.'項目') -join ',' })
    $actual = @((Get-ErrorRows $wb) | ForEach-Object { $_.Key })
    Test-That "エラー確認用:エラーが $($expected.Count) 件と出る" ($message -like "*エラーが $($expected.Count) 件*") $message
    Test-That 'エラー確認用:エラー一覧が期待どおり(シート・行・請求番号・項目、並び順も)' `
        (($actual -join "`n") -eq ($expected -join "`n")) ("実際:" + ($actual -join ' / '))
    $errorSample = Join-Path $work 'error-list-sample.pdf'
    $wb.Worksheets.Item('エラー一覧').ExportAsFixedFormat(0, $errorSample)   # 見本(docs\見本 に置く用)
    $message = Invoke-Macro $wb 'modMain.ExecuteBatch' @($true)
    $history = Get-History $wb
    Test-That 'エラー確認用:PDF作成でも止まり、1件も作らない' ($message -like '*PDFは1件も作成していません*' -and $history[-1].Result -eq 'エラーあり' -and $history[-1].Count -eq '0') $message
    $wb.Close($false)

    #---------------------------------------------------------------------------
    # 正常なデータ
    #---------------------------------------------------------------------------
    $out = Join-Path $work 'out'
    New-Item -ItemType Directory -Path $out | Out-Null
    $wb = Open-Copy $normalBook 'normal.xlsm'
    Set-Setting $wb '保存先フォルダ' $out
    $invoices = @(Import-Csv -Path (Join-Path $root 'testdata\正常\請求一覧.csv') -Encoding UTF8)

    # 入力チェックだけ
    $message = Invoke-Macro $wb 'modMain.ExecuteBatch' @($false)
    Test-That "正常:エラーなし(請求 $($invoices.Count) 件)" ($message -like "エラーはありません。請求 $($invoices.Count) 件*") $message

    # ひな形の計算(端数処理の3つの設定で、全件)
    $tpl = $wb.Worksheets.Item('請求書ひな形')
    foreach ($mode in '切り捨て', '四捨五入', '切り上げ') {
        Set-Setting $wb '消費税の端数処理' $mode
        $expectedTotals = Get-ExpectedTotals $mode
        $bad = @()
        foreach ($inv in $invoices) {
            $no = $inv.'請求番号'
            $msg = Invoke-Macro $wb 'modMain.ExecutePreview' @($no)
            $e = $expectedTotals[$no]
            $got = @([decimal]$tpl.Range('F27').Value2, [decimal]$tpl.Range('F28').Value2, [decimal]$tpl.Range('F29').Value2,
                     [decimal]$tpl.Range('F30').Value2, [decimal]$tpl.Range('F31').Value2, [decimal]$tpl.Range('C8').Value2)
            $want = @($e.S10, $e.T10, $e.S8, $e.T8, $e.Total, $e.Total)
            if (($got -join ',') -ne ($want -join ',') -or $tpl.Range('F4').Text -ne $no -or $tpl.Range('B4').Text -ne $inv.'顧客名') {
                $bad += "$no 実際[$($got -join ',')] 期待[$($want -join ',')] $msg"
            }
        }
        Test-That "ひな形:小計・消費税・合計が合う(端数処理=$mode、$($invoices.Count) 件)" ($bad.Count -eq 0) ($bad -join ' / ')
    }
    Set-Setting $wb '消費税の端数処理' '切り捨て'
    $company = @(7..11 | ForEach-Object { $tpl.Cells.Item($_, 5).Text })
    $wantCompany = @('みほん堂ステーショナリー株式会社', '〒000-0000', '東京都架空区見本町1-2-3 サンプルビル4階', 'TEL 00-0000-0000', '登録番号 T0000000000000')
    Test-That 'ひな形:自社情報(自社名・郵便番号・住所・電話・登録番号)が1行ずつ出る' (($company -join '|') -eq ($wantCompany -join '|')) ($company -join ' / ')
    $msg = Invoke-Macro $wb 'modMain.ExecutePreview' @('INV-9999')
    Test-That 'ひな形:一覧にない請求番号は表示しない' ($msg -like '*表示できません*') $msg

    # PDF作成
    $message = Invoke-Macro $wb 'modMain.ExecuteBatch' @($true)
    $pdfs = @(Get-ChildItem -Path $out -Filter *.pdf)
    $wantNames = @($invoices | ForEach-Object { $_.'請求番号' + '_' + $_.'顧客名' + '.pdf' } | Sort-Object)
    $gotNames = @($pdfs | ForEach-Object { $_.Name } | Sort-Object)
    Test-That "PDF作成:$($invoices.Count) 件と出る" ($message -like "$($invoices.Count) 件の請求書PDFを作成しました*") $message
    Test-That 'PDF作成:ファイル名がルールどおり({請求番号}_{顧客名}.pdf)' (($gotNames -join '|') -eq ($wantNames -join '|')) ($gotNames -join ' / ')
    $infos = @($pdfs | ForEach-Object { Get-PdfInfo $_.FullName })
    Test-That 'PDF作成:どれもPDFで1ページ' (@($infos | Where-Object { -not $_.IsPdf -or $_.Pages -ne 1 }).Count -eq 0) (($infos | ForEach-Object { $_.Pages }) -join ',')
    Test-That 'PDF作成:PDFに作成者(Author)を入れていない' (@($infos | Where-Object { $_.Author }).Count -eq 0)
    Test-That 'PDF作成:終わったあと、ひな形を空に戻す' ($tpl.Range('F4').Text -eq '' -and $tpl.Range('B14').Text -eq '')
    $history = Get-History $wb
    Test-That 'PDF作成:実行履歴に日時・件数・保存先' ($history[-1].Operation -eq 'PDF作成' -and $history[-1].Result -eq '完了' -and $history[-1].Count -eq "$($invoices.Count)" -and $history[-1].Folder -eq $out) ($history[-1] | Out-String)
    $sample = Join-Path $work 'sample.pdf'
    Copy-Item -LiteralPath ($pdfs | Where-Object { $_.Name -like 'INV-1009_*' }).FullName -Destination $sample

    # 同名ファイル:止める
    $stamps = @{}
    foreach ($p in $pdfs) { $stamps[$p.Name] = (Get-Item -LiteralPath $p.FullName).LastWriteTimeUtc }
    $message = Invoke-Macro $wb 'modMain.ExecuteBatch' @($true)
    $errs = Get-ErrorRows $wb
    $changed = @($pdfs | Where-Object { (Get-Item -LiteralPath $_.FullName).LastWriteTimeUtc -ne $stamps[$_.Name] })
    Test-That '同名ファイル(止める):全件エラーにして、1件も書き換えない' ($errs.Count -eq $invoices.Count -and @($errs | Where-Object { $_.Message -like '*同じ名前のファイルがあります*' }).Count -eq $invoices.Count -and $changed.Count -eq 0) $message

    # エラーが1件でもあれば、1件も作らない(空のフォルダで、明細の税率を1つだけ誤りにする)
    $out2 = Join-Path $work 'out2'
    New-Item -ItemType Directory -Path $out2 | Out-Null
    Set-Setting $wb '保存先フォルダ' $out2
    $details = $wb.Worksheets.Item('明細')
    [BatchInvoicePdf.Build.Com]::SetValue2($details.Range('E2'), 0.05)
    $message = Invoke-Macro $wb 'modMain.ExecuteBatch' @($true)
    $errs = Get-ErrorRows $wb
    Test-That 'エラー1件:1件も作らずに止まる' ($errs.Count -eq 1 -and $errs[0].Key -eq '明細,2,INV-1001,税率' -and @(Get-ChildItem $out2).Count -eq 0) ("$message / " + (($errs | ForEach-Object { $_.Key }) -join ' / '))
    [BatchInvoicePdf.Build.Com]::SetValue2($details.Range('E2'), 0.1)
    Set-Setting $wb '保存先フォルダ' $out

    # 想定外のエラー:5件目のPDFを開いたまま(書き込めない)で上書き
    Set-Setting $wb '同名ファイルがあるとき' '上書き'
    $locked = ($pdfs | Where-Object { $_.Name -like 'INV-1005_*' }).FullName
    $lock = [IO.File]::Open($locked, 'Open', 'Read', 'None')
    try {
        $message = Invoke-Macro $wb 'modMain.ExecuteBatch' @($true)
    } finally { $lock.Close() }
    $history = Get-History $wb
    Test-That '想定外のエラー:画面の文に、止まった請求番号と作成件数が出る' ($message -like '*想定外のエラーで止まりました*' -and $message -like '*請求番号:INV-1005*' -and $message -like '*それまでに作成したPDF:4 件*') $message
    Test-That '想定外のエラー:実行履歴に「中断」と請求番号が残る' ($history[-1].Result -eq '中断' -and $history[-1].Count -eq '4' -and $history[-1].Detail -like '*請求番号:INV-1005*') ($history[-1] | Out-String)

    # 同名ファイル:上書き
    $message = Invoke-Macro $wb 'modMain.ExecuteBatch' @($true)
    $pdfs = @(Get-ChildItem -Path $out -Filter *.pdf)
    Test-That '同名ファイル(上書き):全件を作り直す' ($message -like "$($invoices.Count) 件の請求書PDFを作成しました*" -and $pdfs.Count -eq $invoices.Count) $message

    # 設定の誤り
    Set-Setting $wb 'ファイル名ルール' '{請求番号}_{顧客}.pdf'
    Set-Setting $wb '登録番号' 'T123'
    Set-Setting $wb '消費税の端数処理' 'まるめ'
    Set-Setting $wb '同名ファイルがあるとき' '消す'
    $message = Invoke-Macro $wb 'modMain.ExecuteBatch' @($false)
    $keys = @((Get-ErrorRows $wb) | ForEach-Object { ($_.Key -split ',')[3] })
    $wantKeys = (@('ファイル名ルール', '登録番号', '消費税の端数処理', '同名ファイルがあるとき') | Sort-Object) -join ','
    Test-That '設定の誤り:差し込み・登録番号・端数処理・同名ファイルを指摘する' ((($keys | Sort-Object) -join ',') -eq $wantKeys) ($keys -join ',')
    Set-Setting $wb 'ファイル名ルール' '{請求番号}_{顧客名}.pdf'
    Set-Setting $wb '登録番号' 'T0000000000000'
    Set-Setting $wb '消費税の端数処理' '切り捨て'
    Set-Setting $wb '同名ファイルがあるとき' '止める'

    # 相対パス(ブックのある場所から)
    New-Item -ItemType Directory -Path (Join-Path $work '出力') | Out-Null
    Set-Setting $wb '保存先フォルダ' '出力'
    $message = Invoke-Macro $wb 'modMain.ExecuteBatch' @($false)
    $history = Get-History $wb
    Test-That '保存先:相対パス「出力」をブックの場所から探す' ($message -like 'エラーはありません*' -and $history[-1].Folder -eq (Join-Path $work '出力')) "$message / $($history[-1].Folder)"

    # ひな形をコピーして使い分ける(名前の定義もシートと一緒にコピーされる)
    $tpl.Copy([Type]::Missing, $tpl)
    $copied = $wb.Worksheets.Item($tpl.Index + 1)
    Set-Setting $wb 'ひな形シート名' $copied.Name
    $msg = Invoke-Macro $wb 'modMain.ExecutePreview' @('INV-1001')
    Test-That "ひな形:コピーしたシート($($copied.Name))を設定で選べる" ($msg -like "*「$($copied.Name)」シートに表示しました*" -and $copied.Range('F4').Text -eq 'INV-1001') $msg
    Set-Setting $wb 'ひな形シート名' '請求書ひな形'

    # 明細欄を広げる:13行目の明細を足すと止まり、ひな形の途中に1行挿入すると通る
    $details = $wb.Worksheets.Item('明細')
    $next = $details.Cells.Item($details.Rows.Count, 1).End(-4162).Row + 1   # -4162 = 上へ
    $values = @('INV-1016', '追加の品', 1, 100, 0.1)
    for ($c = 0; $c -lt $values.Count; $c++) { [BatchInvoicePdf.Build.Com]::SetValue2($details.Cells.Item($next, $c + 1), $values[$c]) }
    $message = Invoke-Macro $wb 'modMain.ExecuteBatch' @($false)
    $errs = Get-ErrorRows $wb
    Test-That '明細欄:13行の請求はひな形(12行)に入らず止まる' ($errs.Count -eq 1 -and $errs[0].Message -like '*明細が 13 行あり*') $message
    [void]$tpl.Rows.Item(20).Insert()
    $tpl.Range('F20').Formula = '=IF(B20="","",ROUNDDOWN(C20*D20,0))'
    $message = Invoke-Macro $wb 'modMain.ExecuteBatch' @($false)
    $msg = Invoke-Macro $wb 'modMain.ExecutePreview' @('INV-1016')
    Test-That '明細欄:ひな形の途中に1行挿入すると、名前と式が広がって通る' ($message -like 'エラーはありません*' -and $tpl.Range('B26').Text -eq '追加の品' -and $tpl.Range('F32').Text -ne '') "$message / $msg / B26=$($tpl.Range('B26').Text)"

    # 実行後の保存:閉じて開き直しても履歴が残る
    $history = Get-History $wb
    $count = $history.Count
    $wb.Close($false)
    $wb = $app.Workbooks.Open((Join-Path $work 'normal.xlsm'))
    Test-That "実行履歴:閉じて開き直しても残る($count 行)" ((Get-History $wb).Count -eq $count) "$((Get-History $wb).Count) 行"
    $wb.Close($false)
} finally {
    Stop-ExcelSession $session
}

$ng = @($results | Where-Object { -not $_.Ok })
Write-Host ''
Write-Host "結果:$($results.Count) 項目中 OK $($results.Count - $ng.Count) / NG $($ng.Count)"
Write-Host "見本のPDF:$sample / $errorSample"
if (-not $KeepWork -and $ng.Count -eq 0) {
    Get-ChildItem -Path $work -Exclude 'sample.pdf', 'error-list-sample.pdf' | Remove-Item -Recurse -Force
}
Write-Host "作業フォルダ:$work"
if ($ng.Count -gt 0) { exit 1 }
