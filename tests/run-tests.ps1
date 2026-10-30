<#
.SYNOPSIS
    できた .xlsm を Excel で動かして、決まりどおりに動くかを確かめます。

.DESCRIPTION
    リポジトリの .xlsm はそのままにして、一時フォルダ(%TEMP%)にコピーしたものを使います。
    マクロは画面を出さない本体(modMain.ExecuteBatch / ExecutePreview / OutputFolderToCreate、
    modPdf.CreateFolder)を Application.Run で呼びます。画面を出す Run... は呼びません(ボタンを押したときの
    画面は、手で確かめます)。
    確かめること:
      ・エラー確認用のデータで、決めたとおりのエラーが「エラー一覧」に出る(tests\expected-errors.csv)
      ・正常なデータで、エラーがない/20件のPDFができる/ファイル名が決まりどおり/1ページずつ
      ・ひな形の小計・消費税・合計が、CSV から計算した値と合う(端数処理の3つの設定すべて)
      ・エラーが1件でもあれば、1件も作らない
      ・同名ファイル:「止める」で止まり、「上書き」で置き換わる
      ・想定外のエラー(PDFを開いたまま上書き)で、止まった請求番号が画面の文と実行履歴に出る
      ・実行履歴が残り、ブックを閉じて開き直しても消えない。フィルターをかけていても読み落とさない・上書きしない
      ・ボタンにマクロがついている。配る状態のブックに、作成者名や実行の跡が残っていない
      ・エラー値(#N/A)・文字として入った数字・非表示のひな形・長すぎるパス・データなし を、作る前に止める
      ・保存先フォルダがないとき:1つ上があれば作るかを聞く対象になり、[いいえ] ならエラー一覧へ、[はい] なら
        最後の1階層だけ作って続く。1つ上もなければ(ドライブ名の打ち間違いなど)聞かず・作らずに止める

.PARAMETER SampleDir
    指定すると、見本のPDF(請求書 INV-1009 とエラー一覧の例)をそのフォルダに置きます(docs\見本 を作り直すとき)。

.EXAMPLE
    powershell -NoProfile -ExecutionPolicy Bypass -File tests\run-tests.ps1
    powershell -NoProfile -ExecutionPolicy Bypass -File tests\run-tests.ps1 -SampleDir docs\見本
#>
param(
    [string]$BookDir = (Split-Path -Parent $PSScriptRoot),
    [string]$SampleDir = '',
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
    Test-That 'ひな形:長い顧客名・品名は縮小して収める' ($tpl.Range('B4').ShrinkToFit -and $tpl.Range('B14').ShrinkToFit -and $tpl.Range('B25').ShrinkToFit)
    $honorific = ''
    try { $honorific = $wb.Worksheets.Item('請求一覧').Range('D2').Validation.Formula1 } catch { }
    Test-That '請求一覧:敬称をドロップダウンで選べる(御中・様)' ($honorific -eq '御中,様') $honorific
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

    # 請求一覧にフィルターをかけて、下の方の行を隠しても、全件を読む
    $invoiceTable = $wb.Worksheets.Item('請求一覧').ListObjects.Item(1)
    [void]$invoiceTable.Range.AutoFilter(3, '株式会社ひな形システムズ')
    $message = Invoke-Macro $wb 'modMain.ExecuteBatch' @($false)
    $invoiceTable.AutoFilter.ShowAllData()
    Test-That "フィルター:請求一覧の行を隠しても $($invoices.Count) 件とも読む" ($message -like "エラーはありません。請求 $($invoices.Count) 件*") $message

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
    # 見本(値引きの行と 8% の品目がある INV-1009)。-SampleDir のときに docs\見本 などへ置く
    $samplePdf = $pdfs | Where-Object { $_.Name -like 'INV-1009_*' }
    $sampleName = $samplePdf.Name
    $sample = Join-Path $work 'sample.pdf'
    Copy-Item -LiteralPath $samplePdf.FullName -Destination $sample

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

    # 保存先フォルダがないとき。ボタン(Run...)は「このフォルダを作りますか?」と聞いてから、画面を出さない
    # OutputFolderToCreate と modPdf.CreateFolder を呼ぶ。ここでは聞く画面の代わりに、その2つを直接呼ぶ
    $newFolder = Join-Path $work '新しい出力'
    Set-Setting $wb '保存先フォルダ' '新しい出力'   # ブックの場所(ある)の下の、まだないフォルダ
    $ask = [string](Invoke-Macro $wb 'modMain.OutputFolderToCreate')
    Test-That '保存先がない:1つ上のフォルダがあれば、作るかを聞く(フルパスで)' ($ask -eq $newFolder) $ask
    $message = Invoke-Macro $wb 'modMain.ExecuteBatch' @($false)   # [いいえ]:作らずに入力チェックへ
    $errs = Get-ErrorRows $wb
    Test-That '保存先がない:[いいえ] なら作らず、今までどおりエラー一覧に出して止まる' `
        ($errs.Count -eq 1 -and $errs[0].Key -like '設定,*,,保存先フォルダ' -and $errs[0].Message.StartsWith('保存先フォルダがありません:') -and -not (Test-Path -LiteralPath $newFolder)) `
        "$message / $(@($errs | ForEach-Object { $_.Message }) -join ' / ')"
    $reason = [string](Invoke-Macro $wb 'modPdf.CreateFolder' @($ask))   # [はい]:作って、そのまま続ける
    $message = Invoke-Macro $wb 'modMain.ExecuteBatch' @($false)
    Test-That '保存先がない:[はい] ならフォルダを作り、そのまま続けられる' `
        ($reason -eq '' -and (Test-Path -LiteralPath $newFolder -PathType Container) -and $message -like 'エラーはありません*') "$reason / $message"
    $ask = [string](Invoke-Macro $wb 'modMain.OutputFolderToCreate')
    $reason = [string](Invoke-Macro $wb 'modPdf.CreateFolder' @($newFolder))
    Test-That '保存先がある:作るかを聞かない(作る手続きを呼んでも何もしない)' ($ask -eq '' -and $reason -eq '') "$ask / $reason"

    # 1つ上のフォルダもないとき(途中のフォルダもない・ドライブ名の打ち間違い):聞かず、作らず、今までどおり止める
    $drives = @([IO.DriveInfo]::GetDrives() | ForEach-Object { $_.Name.Substring(0, 1).ToUpper() })
    $noDrive = @(68..90 | ForEach-Object { [string][char]$_ } | Where-Object { $drives -notcontains $_ })[-1]   # このPCにないドライブ
    foreach ($case in @(@('途中のフォルダもない', (Join-Path $work 'ない\出力')), @('このPCにないドライブ', "$($noDrive):\請求書PDF"))) {
        $bad = $case[1]
        Set-Setting $wb '保存先フォルダ' $bad
        $ask = [string](Invoke-Macro $wb 'modMain.OutputFolderToCreate')
        $reason = [string](Invoke-Macro $wb 'modPdf.CreateFolder' @($bad))
        $message = Invoke-Macro $wb 'modMain.ExecuteBatch' @($false)
        $errs = Get-ErrorRows $wb
        Test-That "保存先の1つ上もない($($case[0])):聞かず、作らず、エラー一覧に出して止まる" `
            ($ask -eq '' -and $reason -ne '' -and -not (Test-Path -LiteralPath (Split-Path -Parent $bad)) -and
             $errs.Count -eq 1 -and $errs[0].Message.StartsWith('保存先フォルダも、その1つ上のフォルダもありません:')) `
            "$bad / $ask / $reason / $(@($errs | ForEach-Object { $_.Message }) -join ' / ')"
    }

    # ドライブのすぐ下(「C:\請求書PDF」の形)は、1つ上がドライブ(「C:\」)でも作るかを聞く(ここでは作らない)
    $underDrive = (Split-Path -Qualifier $work) + '\batch-invoice-pdf-test-' + [Guid]::NewGuid().ToString('N').Substring(0, 8)
    Set-Setting $wb '保存先フォルダ' $underDrive
    $ask = [string](Invoke-Macro $wb 'modMain.OutputFolderToCreate')
    Test-That '保存先がドライブのすぐ下(C:\請求書PDF の形):作るかを聞く' ($ask -eq $underDrive -and -not (Test-Path -LiteralPath $underDrive)) $ask

    # 作れないとき(同じ名前のファイルがある など)は、エラーで止まらずに理由を返す(ボタンからなら、その理由を見せる)
    $blocker = Join-Path $work '同じ名前のファイル'
    Set-Content -LiteralPath $blocker -Value 'x'
    $reason = [string](Invoke-Macro $wb 'modPdf.CreateFolder' @($blocker))
    Test-That '作れないとき:エラーで止まらず、理由を返す' ($reason -ne '' -and (Test-Path -LiteralPath $blocker -PathType Leaf)) $reason
    Set-Setting $wb '保存先フォルダ' '出力'

    # 実行履歴にフィルターをかけて、いちばん下の行を隠しても、その下に足す(隠れた行を上書きしない)
    $historySheet = $wb.Worksheets.Item('実行履歴')
    $before = Get-History $wb
    [void]$historySheet.Range("A3:F$(3 + $before.Count)").AutoFilter(2, 'PDF作成')   # 「入力チェック」の行を隠す(いちばん下の行も)
    $message = Invoke-Macro $wb 'modMain.ExecuteBatch' @($false)
    $historySheet.AutoFilterMode = $false
    $after = Get-History $wb
    Test-That 'フィルター:実行履歴の行を隠しても、隠れた行を上書きせず下に足す' `
        ($after.Count -eq $before.Count + 1 -and $after[-2].Operation -eq $before[-1].Operation -and $after[-2].Detail -eq $before[-1].Detail -and $after[-1].Operation -eq '入力チェック') `
        "前 $($before.Count) 行 → 後 $($after.Count) 行"

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

    #---------------------------------------------------------------------------
    # 作る前に止めるもの:エラー値・文字の数字・非表示のひな形・長すぎるパス・データなし
    #---------------------------------------------------------------------------
    $outEdge = Join-Path $work 'out-edge'
    New-Item -ItemType Directory -Path $outEdge | Out-Null
    $wb = Open-Copy $normalBook 'edge.xlsm'
    Set-Setting $wb '保存先フォルダ' $outEdge
    $invoiceSheet = $wb.Worksheets.Item('請求一覧')
    $details = $wb.Worksheets.Item('明細')
    $invoiceSheet.Range('C2').Formula = '=NA()'                          # 顧客名がエラー値(VLOOKUP の失敗など)
    $details.Range('C2').NumberFormat = '@'
    [BatchInvoicePdf.Build.Com]::SetValue2($details.Range('C2'), '5')    # 数量が文字の「5」
    $wb.Worksheets.Item('請求書ひな形').Visible = 0                        # ひな形を非表示
    $message = Invoke-Macro $wb 'modMain.ExecuteBatch' @($true)
    $errs = Get-ErrorRows $wb
    $pick = { param($key, $like) @($errs | Where-Object { $_.Key -eq $key -and $_.Message -like $like }).Count }
    Test-That 'エラー値:顧客名の #N/A を「エラー値」と知らせる' ((& $pick '請求一覧,2,INV-1001,顧客名' '*エラー値*') -eq 1) ($errs | ForEach-Object { $_.Key } | Select-Object -First 5 | Out-String)
    Test-That '文字の数字:数量の「5」(文字)を「文字として入っています」と知らせる' ((& $pick '明細,2,INV-1001,数量' '*文字として*') -eq 1)
    Test-That '非表示のひな形:PDFにできないと、作る前に知らせる' ((& $pick '請求書ひな形,,,シートの表示' '*非表示*') -eq 1)
    Test-That '上の3つ:1件も作らずに止まる' ($message -like '*PDFは1件も作成していません*' -and @(Get-ChildItem $outEdge).Count -eq 0) $message
    [BatchInvoicePdf.Build.Com]::SetValue2($invoiceSheet.Range('C2'), '株式会社ひな形システムズ')
    [void]$details.Range('C2').ClearFormats()
    [BatchInvoicePdf.Build.Com]::SetValue2($details.Range('C2'), 5.0)
    $wb.Worksheets.Item('請求書ひな形').Visible = -1

    # 顧客名は式で出してもよい(手順書に書いたこと):式の結果が請求書に出る
    $invoiceSheet.Range('C3').Formula = '="例題建設"&"株式会社"'
    $msg = Invoke-Macro $wb 'modMain.ExecutePreview' @('INV-1002')
    Test-That '式:顧客名を式で出しても、その結果が請求書に出る' ($wb.Worksheets.Item('請求書ひな形').Range('B4').Text -eq '例題建設株式会社') "$msg / B4=$($wb.Worksheets.Item('請求書ひな形').Range('B4').Text)"
    [BatchInvoicePdf.Build.Com]::SetValue2($invoiceSheet.Range('C3'), '例題建設株式会社')

    # 保存先とファイル名を合わせて 255 文字を超えるものだけを止める(Excel のPDF保存の上限)
    $long = $work
    while ($long.Length -lt 230) { $long = Join-Path $long ('d' * [Math]::Min(50, 232 - $long.Length - 1)) }
    New-Item -ItemType Directory -Path $long -Force | Out-Null
    Set-Setting $wb '保存先フォルダ' $long
    $message = Invoke-Macro $wb 'modMain.ExecuteBatch' @($false)
    $tooLong = @($invoices | Where-Object { ($long + '\' + $_.'請求番号' + '_' + $_.'顧客名' + '.pdf').Length -gt 255 })
    $errs = Get-ErrorRows $wb
    $longErrs = @($errs | Where-Object { $_.Message -like '*255 文字を超えます*' })
    Test-That "パスの長さ:255 文字を超える $($tooLong.Count) 件だけを止める(保存先 $($long.Length) 文字)" `
        ($tooLong.Count -gt 0 -and $tooLong.Count -lt $invoices.Count -and $longErrs.Count -eq $tooLong.Count -and $errs.Count -eq $tooLong.Count) `
        "期待 $($tooLong.Count) 件 / エラー $($errs.Count) 件 / $message"

    # 差し込み先のセルのロックを外せば、ひな形を保護してもPDFを作れる(引き継ぎメモに書いたこと)
    $outProtect = Join-Path $work 'out-protect'
    New-Item -ItemType Directory -Path $outProtect | Out-Null
    Set-Setting $wb '保存先フォルダ' $outProtect
    $protectedTemplate = $wb.Worksheets.Item('請求書ひな形')
    foreach ($n in '差込_請求番号', '差込_請求日', '差込_顧客名', '差込_敬称', '差込_支払期限', '差込_品名', '差込_数量', '差込_単価', '差込_税率') {
        foreach ($cell in $protectedTemplate.Range($n).Cells) { $cell.MergeArea.Locked = $false }
    }
    $protectedTemplate.Protect()
    $message = Invoke-Macro $wb 'modMain.ExecuteBatch' @($true)
    $protectedTemplate.Unprotect()
    Test-That '保護したひな形:差し込み先のロックを外せば、全件PDFにできる' `
        ($message -like "$($invoices.Count) 件の請求書PDFを作成しました*" -and @(Get-ChildItem $outProtect -Filter *.pdf).Count -eq $invoices.Count) $message

    # データが1件もないときは、0 件で「完了」にしない
    Set-Setting $wb '保存先フォルダ' $outEdge
    [void]$invoiceSheet.ListObjects.Item(1).DataBodyRange.ClearContents()
    $message = Invoke-Macro $wb 'modMain.ExecuteBatch' @($true)
    $errs = Get-ErrorRows $wb
    Test-That 'データなし:「請求が1件もありません」で止まる' ((& $pick '請求一覧,,,データ' '*1件もありません*') -eq 1 -and $message -like '*PDFは1件も作成していません*') $message
    $wb.Close($false)
} finally {
    Stop-ExcelSession $session
}

$ng = @($results | Where-Object { -not $_.Ok })
Write-Host ''
Write-Host "結果:$($results.Count) 項目中 OK $($results.Count - $ng.Count) / NG $($ng.Count)"
if ($ng.Count -eq 0 -and $SampleDir) {
    $dest = $(if ([IO.Path]::IsPathRooted($SampleDir)) { $SampleDir } else { Join-Path $root $SampleDir })
    New-Item -ItemType Directory -Force -Path $dest | Out-Null
    Copy-Item -LiteralPath $sample -Destination (Join-Path $dest $sampleName) -Force
    Copy-Item -LiteralPath $errorSample -Destination (Join-Path $dest 'エラー一覧の例.pdf') -Force
    Write-Host "見本のPDFを置きました:$dest"
}
if (-not $KeepWork -and $ng.Count -eq 0) {
    Remove-Item -Path $work -Recurse -Force
} else {
    Write-Host "作業フォルダ(確かめたあとに消してかまいません):$work"
}
if ($ng.Count -gt 0) { exit 1 }
