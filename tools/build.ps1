<#
.SYNOPSIS
    src\ の VBA と testdata\ のデータから、請求書一括PDF作成の .xlsm を組み立てます。

.DESCRIPTION
    1. Excel を自動操作して、シート・書式・ひな形・データを作り、いったん .xlsx に保存する
    2. src\ の .bas / .cls から vbaProject.bin を作って入れ、.xlsm にする(tools\VbaProjectWriter.cs)
    3. その .xlsm を Excel で開き、ボタンをつけて保存し直す
    Excel のセキュリティ設定(VBA プロジェクトへのアクセスの信頼など)は使わず、変えもしません。
    途中のファイルは一時フォルダ(%TEMP%)に作り、最後に消します。

.PARAMETER OutDir
    できた .xlsm を置くフォルダ。省略するとリポジトリの直下。

.EXAMPLE
    powershell -NoProfile -ExecutionPolicy Bypass -File tools\build.ps1
#>
param(
    [string]$OutDir = (Split-Path -Parent $PSScriptRoot)
)

$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
. (Join-Path $PSScriptRoot 'ExcelSession.ps1')
Add-Type -AssemblyName System.IO.Compression, System.IO.Compression.FileSystem
if (-not ('BatchInvoicePdf.Build.VbaProjectWriter' -as [type])) {
    Add-Type -Path (Join-Path $PSScriptRoot 'VbaProjectWriter.cs') `
        -ReferencedAssemblies System.IO.Compression, System.IO.Compression.FileSystem, System.Xml
}

#-------------------------------------------------------------------------------
# 作るブック
#-------------------------------------------------------------------------------
$books = @(
    @{ FileName = '請求書一括PDF作成.xlsm';            Data = 'testdata\正常' },
    @{ FileName = '請求書一括PDF作成_エラー確認用.xlsm'; Data = 'testdata\エラー確認用' }
)

# シートの並び順・表示名・コード名(VBA からはコード名でシートを指す)
$sheetDefs = [ordered]@{
    '使い方'       = 'shUsage'
    '設定'         = 'shSettings'
    '請求一覧'     = 'shInvoices'
    '明細'         = 'shDetails'
    '請求書ひな形' = 'shTemplate'
    'エラー一覧'   = 'shErrors'
    '実行履歴'     = 'shHistory'
}

# VBA プロジェクトの ID(毎回同じにしておく)
$projectGuid = '7A0C51E2-3D94-4B6F-8E21-5C9B0D4F6A13'

# 「設定」シートの行。Group は見出しの行。値はすべて架空。
$settingDefs = @(
    @{ Group = 'PDFの出力' },
    @{ Label = 'ひな形シート名'; Value = '請求書ひな形'
       Note = 'PDFにするシートの名前。ひな形を作り替えるときは、シートをコピーして作り、その名前をここに書きます。' },
    @{ Label = '保存先フォルダ'; Value = '出力'
       Note = 'PDFを保存するフォルダ。「C:\請求書PDF」のようなフルパスか、このブックのある場所からの相対パス(「出力」なら、ブックと同じ場所の「出力」フォルダ)。フォルダがないときは、[入力をチェック] [PDFを作成] を押したときに作るかを聞きます(作るのは最後の1階層だけ。1つ上のフォルダもないときは、聞かずに止まります)。' },
    @{ Label = 'ファイル名ルール'; Value = '{請求番号}_{顧客名}.pdf'
       Note = '使える差し込みは {請求番号} {顧客名} {請求日}(20261031 の形)。最後は .pdf にします。' },
    @{ Label = '同名ファイルがあるとき'; Value = '止める'; Choices = '止める,上書き'
       Note = '「止める」:保存先に同じ名前のファイルが1つでもあれば、1件も作らずにエラー一覧へ出します。「上書き」:置き換えます。' },
    @{ Label = '実行後にブックを保存'; Value = 'する'; Choices = 'する,しない'
       Note = '「する」:実行履歴を残すため、実行のあとにこのブックを上書き保存します。' },
    @{ Group = '請求書に載せる自社情報(見本のため、すべて架空)' },
    @{ Label = '自社名'; Value = 'みほん堂ステーショナリー株式会社'; Note = '' },
    @{ Label = '郵便番号'; Value = '〒000-0000'; Note = '' },
    @{ Label = '住所'; Value = '東京都架空区見本町1-2-3 サンプルビル4階'; Note = '' },
    @{ Label = '電話番号'; Value = '00-0000-0000'; Note = '' },
    @{ Label = '登録番号'; Value = 'T0000000000000'
       Note = '適格請求書発行事業者の登録番号(T と数字13桁)。見本のため架空の番号です。' },
    @{ Label = '振込先'; Value = '架空銀行 見本支店 普通 0000000 ミホンドウステーショナリー(カ'; Note = '' },
    @{ Group = '計算' },
    @{ Label = '消費税の端数処理'; Value = '切り捨て'; Choices = '切り捨て,四捨五入,切り上げ'
       Note = '税率ごとの合計に対して、1回だけ端数を処理します。' }
)

# 「使い方」シートの文章(行, 文, 種類)
$usageLines = @(
    @(1,  '請求書一括PDF作成ツール', 'title'),
    @(2,  '「請求一覧」と「明細」から、請求書を1件ずつPDFにします。PCの中だけで動き、データを外へ送りません。', ''),
    @(4,  '使い方', 'head'),
    @(5,  '1. 「設定」シートで、保存先フォルダと自社情報を確かめます。', ''),
    @(6,  '2. 「請求一覧」と「明細」にデータを入れます。2つは請求番号で結びつきます。', ''),
    @(7,  '3. [入力をチェック] を押します。エラーがあれば「エラー一覧」シートに、場所と直し方が出ます。', ''),
    @(8,  '4. [PDFを作成] を押します。エラーが1件でもあれば、1件も作らずに止まります。', ''),
    @(9,  '5. 実行のたびに「実行履歴」シートへ、日時・作成件数・保存先が残ります。', ''),
    @(11, 'ボタン', 'head'),
    @(14, '見た目を確かめたいとき', 'head'),
    @(15, '[1件だけひな形に表示] で、指定した請求を「請求書ひな形」シートに表示します(PDFは作りません)。', ''),
    @(16, '請求番号を入れる欄には、最初は「請求一覧」の先頭の請求番号が入っています。', ''),
    @(18, '止まったとき', 'head'),
    @(19, '・入力の誤り … 「エラー一覧」シートに、シート・行・項目と直し方が出ます。', ''),
    @(20, '・それ以外 … 画面のメッセージと「実行履歴」シートに、止まった段階と請求番号が出ます。', ''),
    @(21, '・くわしくは手順書(docs\手順書.md)を見てください。', ''),
    @(23, 'データの扱い', 'head'),
    @(24, 'このブックは、ネット通信・外部サービス・アドインを使いません。PDFはこのPCのフォルダにだけ保存します。', ''),
    @(25, 'PDFには、ブックの作成者名などの情報を入れません。', '')
)
$buttonRow = 12
$buttonDefs = @(
    @{ Caption = '入力をチェック';      Macro = 'modMain.RunCheckOnly' },
    @{ Caption = 'PDFを作成';           Macro = 'modMain.RunCreatePdfs' },
    @{ Caption = '1件だけひな形に表示'; Macro = 'modMain.RunPreview' }
)

$invoiceHeaders = @('請求番号', '請求日', '顧客名', '敬称', '支払期限')
$detailHeaders = @('請求番号', '品名', '数量', '単価', '税率')

#-------------------------------------------------------------------------------
# 小さな道具
#-------------------------------------------------------------------------------
$colorDark = ConvertTo-ExcelColor 64 64 64
$colorPale = ConvertTo-ExcelColor 242 242 242
$colorAccent = ConvertTo-ExcelColor 31 78 121
$colorWhite = ConvertTo-ExcelColor 255 255 255

# CSV の文字を、セルに入れる値にする(日付は日付の数値、% は小数、数字は数値。読めないものは文字のまま)
# 文字の列(請求番号・顧客名 など)には使わない(「0012」が 12 になってしまうため)
function ConvertTo-CellValue {
    param([string]$Text)
    if ([string]::IsNullOrEmpty($Text)) { return $null }
    $date = [datetime]::MinValue
    if ($Text -match '^\d{4}/\d{1,2}/\d{1,2}$' -and
        [datetime]::TryParseExact($Text, 'yyyy/M/d', [Globalization.CultureInfo]::InvariantCulture, 'None', [ref]$date)) {
        return $date.ToOADate()
    }
    if ($Text -match '^-?\d+(\.\d+)?%$') { return [double]$Text.TrimEnd('%') / 100 }
    if ($Text -match '^-?\d+(\.\d+)?$') { return [double]$Text }
    return $Text
}

# TextHeaders の列は、CSV の文字をそのまま入れる
function Read-Table {
    param([string]$Path, [string[]]$Headers, [string[]]$TextHeaders)
    $rows = @(Import-Csv -Path $Path -Encoding UTF8)
    $data = New-Object 'object[,]' ($rows.Count + 1), $Headers.Count
    for ($c = 0; $c -lt $Headers.Count; $c++) { $data[0, $c] = $Headers[$c] }
    for ($r = 0; $r -lt $rows.Count; $r++) {
        for ($c = 0; $c -lt $Headers.Count; $c++) {
            $text = $rows[$r].($Headers[$c])
            if ($TextHeaders -contains $Headers[$c]) {
                $data[($r + 1), $c] = $(if ([string]::IsNullOrEmpty($text)) { $null } else { $text })
            } else {
                $data[($r + 1), $c] = ConvertTo-CellValue $text
            }
        }
    }
    , $data
}

function Set-HeaderStyle {
    param($Range)
    $Range.Font.Bold = $true
    $Range.Font.Color = $colorWhite
    $Range.Interior.Color = $colorDark
    $Range.HorizontalAlignment = -4108   # 中央
}

function Set-Title {
    param($Range, [int]$Size = 14)
    $Range.Font.Size = $Size
    $Range.Font.Bold = $true
}

function Set-View {
    param($Sheet, [int]$FreezeRow = 0, [bool]$Gridlines = $true)
    try {
        $Sheet.Activate()
        $win = $Sheet.Application.ActiveWindow
        $win.DisplayGridlines = $Gridlines
        $win.FreezePanes = $false
        if ($FreezeRow -gt 0) {
            $win.ScrollRow = 1
            $win.SplitColumn = 0
            $win.SplitRow = $FreezeRow
            $win.FreezePanes = $true
        }
        [void]$Sheet.Range('A1').Select()
    } catch {
        Write-Warning "表示の設定をとばしました($($Sheet.Name)):$($_.Exception.Message)"
    }
}

#-------------------------------------------------------------------------------
# シートを作る
#-------------------------------------------------------------------------------
function Build-Usage {
    param($ws)
    $ws.Range('A:A').ColumnWidth = 110
    foreach ($line in $usageLines) {
        $cell = $ws.Cells.Item($line[0], 1)
        $cell.Value2 = $line[1]
        switch ($line[2]) {
            'title' { Set-Title $cell 16 }
            'head'  { $cell.Font.Bold = $true; $cell.Font.Color = $colorAccent }
        }
    }
    $ws.Rows.Item($buttonRow).RowHeight = 40
}

# RowOf に、項目名 → 行番号 を入れます(ひな形の式が設定のセルを指すのに使う)
function Build-Settings {
    param($ws, [hashtable]$Overrides, [hashtable]$RowOf)
    $ws.Range('A:A').ColumnWidth = 24
    $ws.Range('B:B').ColumnWidth = 46
    $ws.Range('C:C').ColumnWidth = 80
    $ws.Range('A1').Value2 = '設定'
    Set-Title $ws.Range('A1')
    $ws.Range('A2').Value2 = 'ここを書き換えると動きが変わります。A列の項目名は変えないでください(マクロが項目名で探します)。'
    $ws.Range('A4').Value2 = '項目'
    $ws.Range('B4').Value2 = '値'
    $ws.Range('C4').Value2 = '説明'
    Set-HeaderStyle $ws.Range('A4:C4')
    $ws.Range('B:B').NumberFormat = '@'

    $r = 5
    foreach ($def in $settingDefs) {
        if ($def.ContainsKey('Group')) {
            $ws.Cells.Item($r, 1).Value2 = '【' + $def.Group + '】'
            $ws.Range("A${r}:C${r}").Interior.Color = $colorPale
            $ws.Cells.Item($r, 1).Font.Bold = $true
        } else {
            $value = $def.Value
            if ($Overrides.ContainsKey($def.Label)) { $value = $Overrides[$def.Label] }
            $ws.Cells.Item($r, 1).Value2 = $def.Label
            $ws.Cells.Item($r, 2).Value2 = $value
            $ws.Cells.Item($r, 3).Value2 = $def.Note
            if ($def.ContainsKey('Choices')) {
                $v = $ws.Cells.Item($r, 2).Validation
                $v.Delete()
                $v.Add(3, 1, 1, $def.Choices)   # 3=リスト、1=停止
                $v.IgnoreBlank = $true
                $v.InCellDropdown = $true
            }
            $RowOf[$def.Label] = $r
        }
        $r++
    }
    $last = $r - 1
    $ws.Range("A4:C$last").Borders.LineStyle = 1
    $ws.Range("C5:C$last").WrapText = $true
    $ws.Range("A5:C$last").VerticalAlignment = -4160   # 上
    foreach ($row in $RowOf.Values) { $ws.Cells.Item($row, 2).Interior.Color = ConvertTo-ExcelColor 255 255 230 }   # 書き換える欄
}

# TextHeaders:CSV の文字をそのまま入れる列(請求番号・顧客名 など)
# Choices:見出し → 選択肢(カンマ区切り)。その列にドロップダウンをつける(一覧にない値も入れられる)
function Build-DataSheet {
    param($ws, [string]$CsvPath, [string[]]$Headers, [string]$TableName, [hashtable]$Formats, [hashtable]$Widths,
          [string[]]$TextHeaders, [hashtable]$Choices = @{})
    $data = Read-Table -Path $CsvPath -Headers $Headers -TextHeaders $TextHeaders
    $rows = $data.GetLength(0)
    $cols = $data.GetLength(1)
    foreach ($col in $Formats.Keys) { $ws.Range("${col}:${col}").NumberFormat = $Formats[$col] }
    foreach ($col in $Widths.Keys) { $ws.Range("${col}:${col}").ColumnWidth = $Widths[$col] }
    $range = $ws.Range($ws.Cells.Item(1, 1), $ws.Cells.Item($rows, $cols))
    [BatchInvoicePdf.Build.Com]::SetValue2($range, $data)
    $table = $ws.ListObjects.Add(1, $range, [Type]::Missing, 1)
    $table.Name = $TableName
    $table.TableStyle = 'TableStyleLight9'
    foreach ($header in $Choices.Keys) {
        $v = $table.ListColumns.Item($header).DataBodyRange.Validation
        $v.Delete()
        $v.Add(3, 3, 1, $Choices[$header])   # 3=リスト、3=情報
        $v.ShowError = $false                 # 一覧にない値(殿 など)も、メッセージを出さずに入れられる
        $v.InCellDropdown = $true
    }
}

function Build-Template {
    param($ws, [hashtable]$SettingRow)
    $app = $ws.Application
    $sheetRef = "'" + $ws.Name + "'!"
    $setRef = { param($label) "'設定'!`$B`$" + $SettingRow[$label] }

    foreach ($w in @(@('A', 2), @('B', 36), @('C', 9), @('D', 13), @('E', 15), @('F', 16), @('G', 2))) {
        $ws.Range("$($w[0]):$($w[0])").ColumnWidth = $w[1]
    }

    # 題
    $ws.Range('B2:F2').Merge()
    $ws.Range('B2').Value2 = '請　求　書'
    Set-Title $ws.Range('B2') 20
    $ws.Range('B2').HorizontalAlignment = -4108
    $ws.Range('B2:F2').Borders.Item(9).LineStyle = -4119   # 下:二重線

    # 宛名・請求番号・請求日
    $ws.Range('B4:C4').Merge()
    $ws.Range('B4:C4').ShrinkToFit = $true   # 長い顧客名も切れずに1行に収める(結合1つだけなので安全)
    $ws.Range('B4:D4').Font.Size = 14
    $ws.Range('B4:D4').Borders.Item(9).LineStyle = 1
    $ws.Range('B4').NumberFormat = '@'
    $ws.Range('E4').Value2 = '請求番号'
    $ws.Range('E5').Value2 = '請求日'
    $ws.Range('F4').NumberFormat = '@'
    $ws.Range('F4:F5').HorizontalAlignment = -4152   # 右
    $ws.Range('F5').NumberFormat = 'yyyy"年"m"月"d"日"'
    $ws.Range('B6').Value2 = '下記のとおりご請求申し上げます。'

    # 自社情報(設定シートを参照する式)。1行ずつ E:F を結合する。
    # 「縮小して全体を表示」などの配置の設定を、結合した行をまたいでまとめて行うと、
    # Excel がまとめて1つに結合してしまうので、行ごとに設定する。
    $company = @(
        @(7,  ('=' + (& $setRef '自社名'))),
        @(8,  ('=' + (& $setRef '郵便番号'))),
        @(9,  ('=' + (& $setRef '住所'))),
        @(10, ('="TEL "&' + (& $setRef '電話番号'))),
        @(11, ('="登録番号 "&' + (& $setRef '登録番号')))
    )
    foreach ($line in $company) {
        $area = $ws.Range("E$($line[0]):F$($line[0])")
        $area.Merge()
        $area.ShrinkToFit = $true
        $ws.Range("E$($line[0])").Formula = $line[1]
    }
    $ws.Range('E7').Font.Bold = $true
    $ws.Range('E7').Font.Size = 12

    # ご請求金額・支払期限・振込先
    $ws.Range('B8').Value2 = 'ご請求金額(税込)'
    $ws.Range('B8').Font.Bold = $true
    $ws.Range('B8').Interior.Color = $colorPale
    $ws.Range('C8:D8').Merge()
    $ws.Range('C8').Formula = '=F31'
    $ws.Range('C8').NumberFormat = '"¥"#,##0'
    $ws.Range('C8').Font.Size = 16
    $ws.Range('C8').Font.Bold = $true
    $ws.Range('C8').HorizontalAlignment = -4152
    $ws.Range('C8').IndentLevel = 1   # 右の自社情報と離す
    $ws.Range('B8:D8').Borders.Item(9).LineStyle = 1
    $ws.Range('B8:D8').Borders.Item(9).Weight = -4138   # 中太
    $ws.Range('B9').Value2 = 'お支払期限'
    $ws.Range('C9:D9').Merge()
    $ws.Range('C9').NumberFormat = 'yyyy"年"m"月"d"日"'
    $ws.Range('C9').HorizontalAlignment = -4131   # 左
    $ws.Range('B10').Value2 = 'お振込先'
    $ws.Range('C10:D12').Merge()
    $ws.Range('C10').Formula = '=' + (& $setRef '振込先')
    $ws.Range('C10').WrapText = $true
    $ws.Range('C10').VerticalAlignment = -4160
    $ws.Range('C10').Font.Size = 9
    $ws.Range('B8:B10').VerticalAlignment = -4108

    # 明細(14行目から25行目の12行)
    $ws.Range('B13').Value2 = '品名'
    $ws.Range('C13').Value2 = '数量'
    $ws.Range('D13').Value2 = '単価'
    $ws.Range('E13').Value2 = '税率'
    $ws.Range('F13').Value2 = '金額'
    Set-HeaderStyle $ws.Range('B13:F13')
    $ws.Range('B14:B25').NumberFormat = '@'
    $ws.Range('B14:B25').ShrinkToFit = $true   # 長い品名が隣の列で切れないように
    $ws.Range('D14:D25').NumberFormat = '#,##0'
    $ws.Range('E14:E25').NumberFormat = '[=0.08]"8%※";0%'
    $ws.Range('E14:E25').HorizontalAlignment = -4108
    $ws.Range('F14:F25').Formula = '=IF(B14="","",ROUNDDOWN(C14*D14,0))'
    $ws.Range('F14:F25').NumberFormat = '#,##0'
    $ws.Range('14:25').RowHeight = 20
    $ws.Range('B13:F25').Borders.LineStyle = 1
    $ws.Range('B14:F25').Borders.Item(12).Color = ConvertTo-ExcelColor 191 191 191   # 内側の横線は薄く

    # 税率ごとの小計・消費税・合計
    # 設定の値の前後に空白があっても、マクロのチェック(前後の空白を取って比べる)と同じ結果になるよう TRIM する
    $round = { param($expr)
        $r = 'TRIM(' + (& $setRef '消費税の端数処理') + ')'
        "=IF($r=""切り上げ"",ROUNDUP($expr,0),IF($r=""四捨五入"",ROUND($expr,0),ROUNDDOWN($expr,0)))"
    }
    $ws.Range('B27').Value2 = '※は軽減税率(8%)の対象です。'
    $ws.Range('B27').Font.Size = 9
    $ws.Range('E27').Value2 = '10%対象 小計'
    $ws.Range('F27').Formula = '=SUMIF($E$14:$E$25,0.1,$F$14:$F$25)'
    $ws.Range('E28').Value2 = '消費税(10%)'
    $ws.Range('F28').Formula = & $round 'F27*10/100'
    $ws.Range('E29').Value2 = '8%対象 小計'
    $ws.Range('F29').Formula = '=SUMIF($E$14:$E$25,0.08,$F$14:$F$25)'
    $ws.Range('E30').Value2 = '消費税(8%)'
    $ws.Range('F30').Formula = & $round 'F29*8/100'
    $ws.Range('E31').Value2 = '合計(税込)'
    $ws.Range('F31').Formula = '=F27+F28+F29+F30'
    $ws.Range('F27:F31').NumberFormat = '#,##0'
    $ws.Range('E27:F31').Borders.LineStyle = 1
    $ws.Range('E27:E31').Interior.Color = $colorPale
    $ws.Range('E31:F31').Font.Bold = $true
    $ws.Range('E31:F31').Borders.Item(8).LineStyle = -4119   # 上:二重線

    # 差し込み先の名前(このシートだけの名前)
    $names = [ordered]@{
        '差込_請求番号' = '$F$4';  '差込_請求日' = '$F$5'; '差込_顧客名' = '$B$4'; '差込_敬称' = '$D$4'
        '差込_支払期限' = '$C$9';  '差込_品名' = '$B$14:$B$25'; '差込_数量' = '$C$14:$C$25'
        '差込_単価' = '$D$14:$D$25'; '差込_税率' = '$E$14:$E$25'
    }
    foreach ($n in $names.Keys) { [void]$ws.Names.Add($n, '=' + $sheetRef + $names[$n]) }

    # 印刷(A4縦・1ページ)
    $app.PrintCommunication = $false
    $ps = $ws.PageSetup
    $ps.PrintArea = '$A$1:$G$32'
    $ps.PaperSize = 9            # A4
    $ps.Orientation = 1          # 縦
    $ps.Zoom = $false
    $ps.FitToPagesWide = 1
    $ps.FitToPagesTall = 1
    $ps.CenterHorizontally = $true
    $ps.TopMargin = $app.CentimetersToPoints(1.5)
    $ps.BottomMargin = $app.CentimetersToPoints(1.5)
    $ps.LeftMargin = $app.CentimetersToPoints(1.5)
    $ps.RightMargin = $app.CentimetersToPoints(1.5)
    $app.PrintCommunication = $true
}

# 一覧のシートを印刷するとき用(A4横・幅を1ページに・見出しの行を毎ページに)
function Set-ListPrint {
    param($ws, [string]$TitleRows)
    $app = $ws.Application
    $app.PrintCommunication = $false
    $ps = $ws.PageSetup
    $ps.PaperSize = 9
    $ps.Orientation = 2   # 横
    $ps.Zoom = $false
    $ps.FitToPagesWide = 1
    $ps.FitToPagesTall = $false
    $ps.PrintTitleRows = $TitleRows
    $app.PrintCommunication = $true
}

function Build-Errors {
    param($ws)
    $ws.Range('A1').Value2 = 'エラー一覧'
    Set-Title $ws.Range('A1')
    $ws.Range('A2').Value2 = 'まだチェックしていません。'
    $headers = @('No', 'シート', '行', '請求番号', '項目', '内容')
    for ($c = 0; $c -lt $headers.Count; $c++) { $ws.Cells.Item(4, $c + 1).Value2 = $headers[$c] }
    Set-HeaderStyle $ws.Range('A4:F4')
    $widths = @(5, 14, 6, 14, 18, 100)
    for ($c = 0; $c -lt $widths.Count; $c++) { $ws.Columns.Item($c + 1).ColumnWidth = $widths[$c] }
    $ws.Range('D:D').NumberFormat = '@'
    $ws.Range('F:F').WrapText = $true
    Set-ListPrint $ws '$4:$4'
}

function Build-History {
    param($ws)
    $ws.Range('A1').Value2 = '実行履歴'
    Set-Title $ws.Range('A1')
    $ws.Range('A2').Value2 = 'ボタンを押すたびに1行ずつ増えます。止まったときは「内容」に、止まった段階と請求番号が出ます。'
    $headers = @('実行日時', '操作', '結果', '作成件数', '保存先', '内容')
    for ($c = 0; $c -lt $headers.Count; $c++) { $ws.Cells.Item(3, $c + 1).Value2 = $headers[$c] }
    Set-HeaderStyle $ws.Range('A3:F3')
    $widths = @(19, 12, 10, 9, 50, 90)
    for ($c = 0; $c -lt $widths.Count; $c++) { $ws.Columns.Item($c + 1).ColumnWidth = $widths[$c] }
    $ws.Range('A4:A10000').NumberFormat = 'yyyy/mm/dd hh:mm:ss'
    foreach ($word in @('エラーあり', '中断')) {
        $fc = $ws.Range('C4:C10000').FormatConditions.Add(1, 3, "=""$word""")   # セルの値が等しい
        $fc.Font.Color = ConvertTo-ExcelColor 192 0 0
        $fc.Font.Bold = $true
    }
    Set-ListPrint $ws '$3:$3'
}

function Add-Buttons {
    param($ws)
    $cell = $ws.Cells.Item($buttonRow, 1)
    $left = $cell.Left + 4
    foreach ($def in $buttonDefs) {
        $shape = $ws.Shapes.AddShape(5, $left, $cell.Top + 5, 170, 30)   # 5=角丸四角形
        $shape.Name = 'ボタン_' + $def.Caption
        $shape.OnAction = $def.Macro
        $shape.Placement = 3   # セルに合わせて動かさない
        $shape.Fill.ForeColor.RGB = $colorAccent
        $shape.Line.Visible = 0
        $text = $shape.TextFrame2
        $text.TextRange.Text = $def.Caption
        $text.TextRange.Font.Size = 11
        $text.TextRange.Font.Bold = -1
        $text.TextRange.Font.Fill.ForeColor.RGB = $colorWhite
        $text.TextRange.ParagraphFormat.Alignment = 2   # 中央
        $text.VerticalAnchor = 3                       # 上下の中央
        $left += 182
    }
}

#-------------------------------------------------------------------------------
# 組み立て
#-------------------------------------------------------------------------------
$work = Join-Path ([IO.Path]::GetTempPath()) ('batch-invoice-pdf-build-' + [Guid]::NewGuid().ToString('N').Substring(0, 8))
New-Item -ItemType Directory -Path $work | Out-Null

# VBA のモジュール(文書モジュール + src の .bas/.cls)
$modules = New-Object 'System.Collections.Generic.List[BatchInvoicePdf.Build.VbaModule]'
$modules.Add([BatchInvoicePdf.Build.VbaProjectWriter]::DocumentModule('ThisWorkbook', $true))
foreach ($code in $sheetDefs.Values) { $modules.Add([BatchInvoicePdf.Build.VbaProjectWriter]::DocumentModule($code, $false)) }
$documentNames = @('ThisWorkbook') + @($sheetDefs.Values)
foreach ($file in Get-ChildItem -Path (Join-Path $root 'src') -File | Where-Object { $_.Extension -in '.bas', '.cls' } | Sort-Object Name) {
    # シートと ThisWorkbook のモジュールは、上で空のものを作っている。src に置いても使われず、
    # コードが黙って消えるので、ここで止める(コードは標準モジュールかクラスに置く)
    if ($documentNames -contains $file.BaseName) {
        throw "src\$($file.Name) は、シートか ThisWorkbook のモジュールです。この組み立てでは空のモジュールを作るので、コードは標準モジュール(.bas)かクラス(.cls)に置いてください。"
    }
    $modules.Add([BatchInvoicePdf.Build.VbaProjectWriter]::LoadModule($file.FullName, 932))
}
$vbaProject = [BatchInvoicePdf.Build.VbaProjectWriter]::Build($modules, 932, $projectGuid)
Write-Host "VBA モジュール $($modules.Count) 個(vbaProject.bin $($vbaProject.Length) バイト)"

$codeNames = New-Object 'System.Collections.Generic.Dictionary[string,string]'
foreach ($name in $sheetDefs.Keys) { $codeNames[$name] = $sheetDefs[$name] }

$session = Start-ExcelSession -TimeoutSec 600
try {
    $app = $session.App
    foreach ($book in $books) {
        $dataDir = Join-Path $root $book.Data
        $overrides = @{}
        $settingsCsv = Join-Path $dataDir '設定.csv'
        if (Test-Path $settingsCsv) {
            foreach ($row in Import-Csv -Path $settingsCsv -Encoding UTF8) { $overrides[$row.'項目'] = $row.'値' }
        }

        # 1. シート・書式・データを作って .xlsx に保存
        $wb = $app.Workbooks.Add()
        while ($wb.Worksheets.Count -lt $sheetDefs.Count) {
            [void]$wb.Worksheets.Add([Type]::Missing, $wb.Worksheets.Item($wb.Worksheets.Count))
        }
        $i = 1
        foreach ($name in $sheetDefs.Keys) { $wb.Worksheets.Item($i).Name = $name; $i++ }

        Build-Usage $wb.Worksheets.Item('使い方')
        $settingRow = @{}
        Build-Settings $wb.Worksheets.Item('設定') $overrides $settingRow
        # 「文字列」の表示形式は請求番号の列だけ(001 のような番号の 0 が消えないように)。
        # 顧客名・品名などは、式(VLOOKUP など)を入れられるよう、標準のままにする。
        Build-DataSheet $wb.Worksheets.Item('請求一覧') (Join-Path $dataDir '請求一覧.csv') $invoiceHeaders 'T_請求一覧' `
            @{ A = '@'; B = 'yyyy/mm/dd'; E = 'yyyy/mm/dd' } @{ A = 14; B = 12; C = 34; D = 8; E = 12 } `
            @('請求番号', '顧客名', '敬称') @{ '敬称' = '御中,様' }
        Build-DataSheet $wb.Worksheets.Item('明細') (Join-Path $dataDir '明細.csv') $detailHeaders 'T_明細' `
            @{ A = '@'; D = '#,##0'; E = '0%' } @{ A = 14; B = 36; C = 8; D = 12; E = 8 } `
            @('請求番号', '品名')
        Build-Template $wb.Worksheets.Item('請求書ひな形') $settingRow
        Build-Errors $wb.Worksheets.Item('エラー一覧')
        Build-History $wb.Worksheets.Item('実行履歴')

        $xlsx = Join-Path $work 'book.xlsx'
        $wb.SaveAs($xlsx, 51)
        $wb.Close($false)

        # 2. VBA を入れて .xlsm にする
        $xlsm = Join-Path $work 'book.xlsm'
        [BatchInvoicePdf.Build.XlsmPackager]::Inject($xlsx, $xlsm, $vbaProject, 'ThisWorkbook', $codeNames)

        # 3. Excel で開いてボタンをつけ、表示を整えて保存し直す
        $wb = $app.Workbooks.Open($xlsm)
        if (-not $wb.HasVBProject) { throw 'VBA が入っていません(vbaProject.bin を Excel が読めませんでした)。' }
        Add-Buttons $wb.Worksheets.Item('使い方')
        Set-View $wb.Worksheets.Item('実行履歴') -FreezeRow 3
        Set-View $wb.Worksheets.Item('エラー一覧') -FreezeRow 4
        Set-View $wb.Worksheets.Item('請求書ひな形') -Gridlines $false
        Set-View $wb.Worksheets.Item('明細') -FreezeRow 1
        Set-View $wb.Worksheets.Item('請求一覧') -FreezeRow 1
        Set-View $wb.Worksheets.Item('設定') -FreezeRow 4
        Set-View $wb.Worksheets.Item('使い方') -Gridlines $false
        $wb.RemovePersonalInformation = $true   # 作成者名などをファイルに残さない
        $out = Join-Path $OutDir $book.FileName
        $wb.SaveAs($out, 52)
        $wb.Close($false)
        Write-Host "できました:$out"
    }
} finally {
    Stop-ExcelSession $session
    Remove-Item -Path $work -Recurse -Force -ErrorAction SilentlyContinue
}
