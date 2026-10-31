Attribute VB_Name = "modValidate"
'==============================================================================
' modValidate : 入力チェックと「エラー一覧」への書き出し
'------------------------------------------------------------------------------
' ・PDFを作る前に全件を調べます。エラーが1件でもあれば、1件も作りません。
' ・見つけたエラーは clsCheckError にして集め、最後に「エラー一覧」へまとめて書きます。
' ・チェックを足すときは、合う Check... の中に AddError を1つ足します。
'   メッセージには「何が問題か」と「どう直せばよいか」を書いてください。
'==============================================================================
Option Explicit

' セルがエラー値(#N/A など)のときのメッセージ。式(VLOOKUP など)で値を出しているときに起きやすい
Private Const MSG_ERROR_VALUE As String = "セルがエラー値(#N/A など)になっています。式や、式が参照している元のデータを確かめてください。"

' エラーを1件足します。rowNumber が 0 のときは「行」を空にします。
Public Sub AddError(ByVal errorList As Collection, ByVal sheetLabel As String, ByVal rowNumber As Long, _
                    ByVal invoiceNumber As String, ByVal fieldLabel As String, ByVal detail As String)
    Dim e As clsCheckError

    Set e = New clsCheckError
    e.SheetName = sheetLabel
    e.RowNo = rowNumber
    e.InvoiceNo = invoiceNumber
    e.ItemName = fieldLabel
    e.Message = detail
    errorList.Add e
End Sub

' すべてのチェックをします。明細を請求に結びつけ、PDFのファイル名(FileName)もここで決めます。
Public Sub ValidateAll(ByRef st As AppSettings, ByVal invoices As Collection, ByVal allDetails As Collection, _
                       ByVal errorList As Collection)
    Dim capacity As Long
    Dim byNumber As Object

    CheckSettings st, errorList
    capacity = CheckTemplateSheet(st, errorList)
    Set byNumber = CheckInvoices(invoices, errorList)
    CheckDetails allDetails, byNumber, errorList
    CheckDetailCounts invoices, byNumber, capacity, errorList
    CheckFileNames st, invoices, byNumber, errorList
End Sub

' 請求1件(「請求一覧」の1行)を調べます。
Public Sub CheckInvoice(ByVal inv As clsInvoice, ByVal errorList As Collection)
    Dim sheetLabel As String

    sheetLabel = shInvoices.Name
    RequireText errorList, sheetLabel, inv.RowNo, inv.InvoiceNo, COL_INVOICE_NO, inv.InvoiceNo
    CheckDate errorList, sheetLabel, inv.RowNo, inv.InvoiceNo, COL_INVOICE_DATE, inv.InvoiceDate
    RequireText errorList, sheetLabel, inv.RowNo, inv.InvoiceNo, COL_CUSTOMER, inv.CustomerName
    RequireText errorList, sheetLabel, inv.RowNo, inv.InvoiceNo, COL_HONORIFIC, inv.Honorific
    CheckDate errorList, sheetLabel, inv.RowNo, inv.InvoiceNo, COL_DUE_DATE, inv.DueDate

    If IsDateValue(inv.InvoiceDate) And IsDateValue(inv.DueDate) Then
        If CDate(inv.DueDate) < CDate(inv.InvoiceDate) Then
            AddError errorList, sheetLabel, inv.RowNo, inv.InvoiceNo, COL_DUE_DATE, _
                "支払期限が請求日より前になっています。日付を確かめてください。"
        End If
    End If
End Sub

' 明細1行(「明細」の1行)を調べます。税率が正しければ 0.1 か 0.08 にそろえます。
Public Sub CheckDetail(ByVal d As clsDetail, ByVal errorList As Collection)
    Dim sheetLabel As String

    sheetLabel = shDetails.Name
    RequireText errorList, sheetLabel, d.RowNo, d.InvoiceNo, COL_INVOICE_NO, d.InvoiceNo
    CheckDate errorList, sheetLabel, d.RowNo, d.InvoiceNo, COL_TRADE_DATE, d.TradeDate
    RequireText errorList, sheetLabel, d.RowNo, d.InvoiceNo, COL_ITEM, d.ItemName

    If IsBlankValue(d.Quantity) Then
        AddError errorList, sheetLabel, d.RowNo, d.InvoiceNo, COL_QUANTITY, "入っていません(必須です)。"
    ElseIf Not IsNumberValue(d.Quantity) Then
        AddError errorList, sheetLabel, d.RowNo, d.InvoiceNo, COL_QUANTITY, NotNumberMessage(d.Quantity)
    ElseIf d.Quantity <= 0 Then
        AddError errorList, sheetLabel, d.RowNo, d.InvoiceNo, COL_QUANTITY, _
            "0より大きい数にしてください(いまの値:" & CellText(d.Quantity) & ")。"
    End If

    If IsBlankValue(d.UnitPrice) Then
        AddError errorList, sheetLabel, d.RowNo, d.InvoiceNo, COL_UNIT_PRICE, "入っていません(必須です)。"
    ElseIf Not IsNumberValue(d.UnitPrice) Then
        AddError errorList, sheetLabel, d.RowNo, d.InvoiceNo, COL_UNIT_PRICE, NotNumberMessage(d.UnitPrice)
    End If

    If IsBlankValue(d.TaxRate) Then
        AddError errorList, sheetLabel, d.RowNo, d.InvoiceNo, COL_TAX_RATE, "入っていません(必須です)。"
    ElseIf Not IsNumberValue(d.TaxRate) Then
        AddError errorList, sheetLabel, d.RowNo, d.InvoiceNo, COL_TAX_RATE, NotNumberMessage(d.TaxRate)
    ElseIf NormalizeTaxRate(d.TaxRate) = 0 Then
        AddError errorList, sheetLabel, d.RowNo, d.InvoiceNo, COL_TAX_RATE, _
            "10% か 8% にしてください(いまの値:" & TaxRateText(d.TaxRate) & ")。"
    Else
        d.TaxRate = NormalizeTaxRate(d.TaxRate)
    End If
End Sub

' 「エラー一覧」を書き直します(前回の内容は消します)。
Public Sub WriteErrorList(ByVal errorList As Collection, ByVal invoiceCount As Long)
    Dim ws As Worksheet
    Dim data() As Variant
    Dim e As clsCheckError
    Dim lastRow As Long
    Dim i As Long

    Set ws = shErrors
    lastRow = LastValueRow(ws, 1)   ' フィルターで隠れた行も消す
    If lastRow >= ERROR_FIRST_ROW Then
        ws.Range(ws.Cells(ERROR_FIRST_ROW, 1), ws.Cells(lastRow, ERROR_COLUMNS)).ClearContents
    End If

    If errorList.Count = 0 Then
        ws.Range(ERROR_STATUS_CELL).Value = Format$(Now, "yyyy/mm/dd hh:nn:ss") & _
            " のチェック:エラーはありません(請求 " & invoiceCount & " 件)。"
        Exit Sub
    End If

    ReDim data(1 To errorList.Count, 1 To ERROR_COLUMNS)
    For Each e In SortErrors(errorList)
        i = i + 1
        data(i, 1) = i
        data(i, 2) = e.SheetName
        If e.RowNo > 0 Then data(i, 3) = e.RowNo
        data(i, 4) = e.InvoiceNo
        data(i, 5) = e.ItemName
        data(i, 6) = e.Message
    Next e
    ws.Range(ws.Cells(ERROR_FIRST_ROW, 1), ws.Cells(ERROR_FIRST_ROW + errorList.Count - 1, ERROR_COLUMNS)).Value = data
    ws.Range(ERROR_STATUS_CELL).Value = Format$(Now, "yyyy/mm/dd hh:nn:ss") & _
        " のチェック:エラー " & errorList.Count & " 件。上から順に直して、もう一度実行してください。"
End Sub

'------------------------------------------------------------------------------
' ここから下は、ValidateAll の中の各チェック
'------------------------------------------------------------------------------

' 設定の値が、選べる値・正しい形になっているか。保存先フォルダがあるか。
Private Sub CheckSettings(ByRef st As AppSettings, ByVal errorList As Collection)
    Dim sheetLabel As String
    Dim unknown As String

    sheetLabel = shSettings.Name

    If st.IfExists <> "" And st.IfExists <> IF_EXISTS_STOP And st.IfExists <> IF_EXISTS_OVERWRITE Then
        AddError errorList, sheetLabel, SettingRow(SET_IF_EXISTS), "", SET_IF_EXISTS, _
            "「" & IF_EXISTS_STOP & "」か「" & IF_EXISTS_OVERWRITE & "」にしてください(いまの値:" & st.IfExists & ")。"
    End If

    If st.SaveAfterRun <> "" And st.SaveAfterRun <> YES_TEXT And st.SaveAfterRun <> NO_TEXT Then
        AddError errorList, sheetLabel, SettingRow(SET_SAVE_AFTER_RUN), "", SET_SAVE_AFTER_RUN, _
            "「" & YES_TEXT & "」か「" & NO_TEXT & "」にしてください(いまの値:" & st.SaveAfterRun & ")。"
    End If

    If st.TaxRounding <> "" And Not IsOneOf(st.TaxRounding, ROUNDING_CHOICES) Then
        AddError errorList, sheetLabel, SettingRow(SET_TAX_ROUNDING), "", SET_TAX_ROUNDING, _
            "「" & Replace(ROUNDING_CHOICES, ",", "」「") & "」のどれかにしてください(いまの値:" & st.TaxRounding & ")。"
    End If

    If st.RegistrationNo <> "" And Not (st.RegistrationNo Like REGISTRATION_NO_PATTERN) Then
        AddError errorList, sheetLabel, SettingRow(SET_REGISTRATION_NO), "", SET_REGISTRATION_NO, _
            "T のあとに数字13桁で書いてください(例:T0000000000000、いまの値:" & st.RegistrationNo & ")。"
    End If

    If st.FileNameRule <> "" Then
        If LCase$(Right$(st.FileNameRule, 4)) <> ".pdf" Then
            AddError errorList, sheetLabel, SettingRow(SET_FILE_NAME_RULE), "", SET_FILE_NAME_RULE, _
                "最後を「.pdf」にしてください(いまの値:" & st.FileNameRule & ")。"
        End If
        unknown = UnknownTag(st.FileNameRule)
        If unknown <> "" Then
            AddError errorList, sheetLabel, SettingRow(SET_FILE_NAME_RULE), "", SET_FILE_NAME_RULE, _
                "使えない差し込み「" & unknown & "」があります。使えるのは " & _
                TAG_INVOICE_NO & " " & TAG_CUSTOMER & " " & TAG_INVOICE_DATE & " です。"
        End If
    End If

    If st.OutputFolderText <> "" Then
        If st.OutputFolder = "" Then
            AddError errorList, sheetLabel, SettingRow(SET_OUTPUT_FOLDER), "", SET_OUTPUT_FOLDER, _
                "このブックが未保存か、OneDrive などの上で開かれているため、相対パスを使えません。" & _
                "「C:\請求書PDF」のようなフルパスにしてください。"
        ElseIf CanCreateFolder(st.OutputFolder) Then
            ' 1つ上のフォルダはある:ボタンから実行したときに作るかを聞き、[いいえ] だった(または作れなかった)
            AddError errorList, sheetLabel, SettingRow(SET_OUTPUT_FOLDER), "", SET_OUTPUT_FOLDER, _
                "保存先フォルダがありません:" & st.OutputFolder & "。" & _
                "もう一度実行して「このフォルダを作りますか?」に [はい] と答えるか、設定を直してください。"
        ElseIf Not FolderExists(st.OutputFolder) Then
            ' 1つ上のフォルダもないときは、作るかを聞かない(打ち間違いのまま、思わぬ場所に作らないため)
            AddError errorList, sheetLabel, SettingRow(SET_OUTPUT_FOLDER), "", SET_OUTPUT_FOLDER, _
                "保存先フォルダも、その1つ上のフォルダもありません:" & st.OutputFolder & "。" & _
                "ドライブ名やフォルダ名に打ち間違いがないか、設定を確かめてください(正しいときは、フォルダを作ってからもう一度実行します)。"
        End If
    End If
End Sub

' ひな形シートがあり、差し込み先の名前がそろっているか。明細に入る行数を返します(問題があれば 0)。
Private Function CheckTemplateSheet(ByRef st As AppSettings, ByVal errorList As Collection) As Long
    Dim ws As Worksheet

    If st.TemplateSheetName = "" Then Exit Function
    Set ws = FindSheet(st.TemplateSheetName)
    If ws Is Nothing Then
        AddError errorList, shSettings.Name, SettingRow(SET_TEMPLATE_SHEET), "", SET_TEMPLATE_SHEET, _
            "「" & st.TemplateSheetName & "」というシートがありません。シート名を確かめてください。"
        Exit Function
    End If
    CheckTemplateSheet = CheckTemplate(ws, errorList)
End Function

' 請求を1件ずつ調べ、請求番号の重なりも見ます。請求番号 → clsInvoice の辞書を返します。
Private Function CheckInvoices(ByVal invoices As Collection, ByVal errorList As Collection) As Object
    Dim byNumber As Object
    Dim inv As clsInvoice
    Dim first As clsInvoice

    Set byNumber = CreateObject("Scripting.Dictionary")
    byNumber.CompareMode = vbTextCompare   ' 大文字と小文字は同じと見る(ファイル名でも区別されないため)

    For Each inv In invoices
        CheckInvoice inv, errorList
        If inv.InvoiceNo <> "" Then
            If byNumber.Exists(inv.InvoiceNo) Then
                Set first = byNumber(inv.InvoiceNo)
                AddError errorList, shInvoices.Name, inv.RowNo, inv.InvoiceNo, COL_INVOICE_NO, _
                    "請求番号が " & first.RowNo & " 行目と重なっています。請求番号は1件ごとに別にしてください。"
            Else
                byNumber.Add inv.InvoiceNo, inv
            End If
        End If
    Next inv
    Set CheckInvoices = byNumber
End Function

' 明細を1行ずつ調べ、請求に結びつけます。
Private Sub CheckDetails(ByVal allDetails As Collection, ByVal byNumber As Object, ByVal errorList As Collection)
    Dim d As clsDetail

    For Each d In allDetails
        CheckDetail d, errorList
        If d.InvoiceNo <> "" Then
            If byNumber.Exists(d.InvoiceNo) Then
                byNumber(d.InvoiceNo).Details.Add d
            Else
                AddError errorList, shDetails.Name, d.RowNo, d.InvoiceNo, COL_INVOICE_NO, _
                    "この請求番号は「請求一覧」にありません。番号の書き間違いか、請求一覧への入れ忘れです。"
            End If
        End If
    Next d
End Sub

' 明細のない請求、ひな形の明細欄に入りきらない請求がないか。
Private Sub CheckDetailCounts(ByVal invoices As Collection, ByVal byNumber As Object, ByVal capacity As Long, _
                              ByVal errorList As Collection)
    Dim inv As clsInvoice

    For Each inv In invoices
        If IsRegistered(inv, byNumber) Then
            If inv.Details.Count = 0 Then
                AddError errorList, shInvoices.Name, inv.RowNo, inv.InvoiceNo, "明細", _
                    "明細が1行もありません。「明細」シートに、この請求番号の行を入れてください。"
            ElseIf capacity > 0 And inv.Details.Count > capacity Then
                AddError errorList, shInvoices.Name, inv.RowNo, inv.InvoiceNo, "明細", _
                    "明細が " & inv.Details.Count & " 行あり、ひな形の明細欄(" & capacity & " 行)に入りきりません。" & _
                    "請求を分けるか、ひな形の明細欄を広げてください。"
            End If
        End If
    Next inv
End Sub

' ファイル名を決め、使えない文字・重なり・長さ・同名ファイルを調べます。
Private Sub CheckFileNames(ByRef st As AppSettings, ByVal invoices As Collection, ByVal byNumber As Object, _
                           ByVal errorList As Collection)
    Dim usedNames As Object
    Dim inv As clsInvoice
    Dim filePath As String

    ' ルールそのものの誤りは CheckSettings で出しているので、ここでは調べない
    If st.FileNameRule = "" Or UnknownTag(st.FileNameRule) <> "" Then Exit Sub

    Set usedNames = CreateObject("Scripting.Dictionary")
    usedNames.CompareMode = vbTextCompare

    For Each inv In invoices
        If IsRegistered(inv, byNumber) Then
            inv.FileName = BuildFileName(st.FileNameRule, inv)
            If HasInvalidFileChar(inv.FileName) Then
                AddError errorList, shInvoices.Name, inv.RowNo, inv.InvoiceNo, "ファイル名", _
                    "ファイル名に使えない文字(\ / : * ? "" < > | や改行)が入ります:" & inv.FileName & _
                    "。顧客名などを直すか、ファイル名ルールを変えてください。"
            ElseIf usedNames.Exists(inv.FileName) Then
                AddError errorList, shInvoices.Name, inv.RowNo, inv.InvoiceNo, "ファイル名", _
                    "ファイル名が " & usedNames(inv.FileName) & " 行目の請求と同じ「" & inv.FileName & "」になります。" & _
                    "ファイル名ルールに " & TAG_INVOICE_NO & " を入れると重なりません。"
            Else
                usedNames.Add inv.FileName, inv.RowNo
                If st.OutputFolder <> "" Then
                    filePath = st.OutputFolder & "\" & inv.FileName
                    If Len(filePath) > MAX_PATH_LENGTH Then
                        AddError errorList, shInvoices.Name, inv.RowNo, inv.InvoiceNo, "ファイル名", _
                            "保存先とファイル名を合わせた長さが " & MAX_PATH_LENGTH & " 文字を超えます。" & _
                            "保存先を短くするか、ファイル名ルールを短くしてください。"
                    ElseIf st.IfExists = IF_EXISTS_STOP And FileExists(filePath) Then
                        AddError errorList, shInvoices.Name, inv.RowNo, inv.InvoiceNo, "ファイル名", _
                            "保存先に同じ名前のファイルがあります:" & inv.FileName & "。" & _
                            "置き換えてよいなら、設定の「" & SET_IF_EXISTS & "」を「" & IF_EXISTS_OVERWRITE & "」にしてください。"
                    End If
                End If
            End If
        End If
    Next inv
End Sub

'------------------------------------------------------------------------------
' 小さな道具
'------------------------------------------------------------------------------

' エラーを、シートの順(設定 → ひな形 → 請求一覧 → 明細)と行の順に並べます。
' 同じ行のエラーは、見つけた順のままにします。
Private Function SortErrors(ByVal errorList As Collection) As Collection
    Dim ordered() As clsCheckError
    Dim keys() As Double
    Dim e As clsCheckError
    Dim holdError As clsCheckError
    Dim holdKey As Double
    Dim n As Long, i As Long, j As Long

    ReDim ordered(1 To errorList.Count)
    ReDim keys(1 To errorList.Count)
    For Each e In errorList
        n = n + 1
        Set ordered(n) = e
        keys(n) = SheetRank(e.SheetName) * 1000000# + e.RowNo
    Next e

    ' 件数は多くないので、挿入ソート(同じキーは順番が変わらない)
    For i = 2 To n
        Set holdError = ordered(i)
        holdKey = keys(i)
        j = i - 1
        Do While j >= 1
            If keys(j) <= holdKey Then Exit Do
            Set ordered(j + 1) = ordered(j)
            keys(j + 1) = keys(j)
            j = j - 1
        Loop
        Set ordered(j + 1) = holdError
        keys(j + 1) = holdKey
    Next i

    Set SortErrors = New Collection
    For i = 1 To n
        SortErrors.Add ordered(i)
    Next i
End Function

Private Function SheetRank(ByVal sheetTitle As String) As Long
    Select Case sheetTitle
        Case shSettings.Name: SheetRank = 1
        Case shInvoices.Name: SheetRank = 3
        Case shDetails.Name: SheetRank = 4
        Case Else: SheetRank = 2   ' ひな形シート
    End Select
End Function

' 請求番号が辞書に入っている、その請求そのものか(重なった2件目以降・番号が空の行は False)。
Private Function IsRegistered(ByVal inv As clsInvoice, ByVal byNumber As Object) As Boolean
    If inv.InvoiceNo = "" Then Exit Function
    If Not byNumber.Exists(inv.InvoiceNo) Then Exit Function
    IsRegistered = (byNumber(inv.InvoiceNo) Is inv)
End Function

' 文字の項目が空・エラー値なら、エラーにします(エラー値のまま進むと、PDFに「#エラー」と出てしまう)。
Private Sub RequireText(ByVal errorList As Collection, ByVal sheetLabel As String, ByVal rowNumber As Long, _
                        ByVal invoiceNumber As String, ByVal fieldLabel As String, ByVal fieldText As String)
    If fieldText = "" Then
        AddError errorList, sheetLabel, rowNumber, invoiceNumber, fieldLabel, "入っていません(必須です)。"
    ElseIf fieldText = CELL_ERROR_TEXT Then
        AddError errorList, sheetLabel, rowNumber, invoiceNumber, fieldLabel, MSG_ERROR_VALUE
    End If
End Sub

' 日付の項目が空・日付として読めないなら、エラーにします。
Private Sub CheckDate(ByVal errorList As Collection, ByVal sheetLabel As String, ByVal rowNumber As Long, _
                      ByVal invoiceNumber As String, ByVal fieldLabel As String, ByVal cellValue As Variant)
    If IsBlankValue(cellValue) Then
        AddError errorList, sheetLabel, rowNumber, invoiceNumber, fieldLabel, "入っていません(必須です)。"
    ElseIf IsError(cellValue) Then
        AddError errorList, sheetLabel, rowNumber, invoiceNumber, fieldLabel, MSG_ERROR_VALUE
    ElseIf IsNumberValue(cellValue) Then
        ' 日付が数値(46326 など)で入っている:セルの表示形式が日付になっていない
        AddError errorList, sheetLabel, rowNumber, invoiceNumber, fieldLabel, _
            "日付として読めません(いまの値:" & CellText(cellValue) & ")。セルの表示形式を日付にして、日付を確かめてください。"
    ElseIf Not IsDateValue(cellValue) Then
        AddError errorList, sheetLabel, rowNumber, invoiceNumber, fieldLabel, _
            "日付として読めません(いまの値:" & CellText(cellValue) & ")。2026/10/31 のように入れてください。"
    End If
End Sub

' 数値でない値のメッセージ。エラー値・文字として入った数字は、それとわかるように書きます。
Private Function NotNumberMessage(ByVal cellValue As Variant) As String
    Dim shown As String

    If IsError(cellValue) Then
        NotNumberMessage = MSG_ERROR_VALUE
        Exit Function
    End If

    shown = CellText(cellValue)
    If VarType(cellValue) = vbString And IsNumeric(Replace(Replace(shown, "%", ""), ",", "")) Then
        ' 「5」「10%」などが文字として入っている(表示形式が「文字列」のセル、ほかから貼り付けた値 など)
        NotNumberMessage = "数字が文字として入っています(いまの値:" & shown & ")。" & _
                           "セルの表示形式が「文字列」になっていないかを確かめて、数値として入れ直してください。"
    Else
        NotNumberMessage = "数値にしてください(いまの値:" & shown & ")。"
    End If
End Function

Private Function IsBlankValue(ByVal cellValue As Variant) As Boolean
    IsBlankValue = (Trim$(CellText(cellValue)) = "")
End Function

' 税率を 0.1 か 0.08 にそろえて返します。どちらでもなければ 0。
Private Function NormalizeTaxRate(ByVal cellValue As Variant) As Double
    If Not IsNumberValue(cellValue) Then Exit Function
    If Abs(cellValue - TAX_RATE_STANDARD) < 0.0000001 Then
        NormalizeTaxRate = TAX_RATE_STANDARD
    ElseIf Abs(cellValue - TAX_RATE_REDUCED) < 0.0000001 Then
        NormalizeTaxRate = TAX_RATE_REDUCED
    End If
End Function

' 税率の値を、メッセージ用に「5%」のような形にします(1以上の数はそのまま)。
Private Function TaxRateText(ByVal cellValue As Variant) As String
    If IsNumberValue(cellValue) Then
        If Abs(cellValue) < 1 Then
            TaxRateText = CStr(Round(cellValue * 100, 4)) & "%"
        Else
            TaxRateText = CStr(cellValue) & "。" & CStr(cellValue) & "% のつもりなら「" & CStr(cellValue) & "%」と入れます"
        End If
    Else
        TaxRateText = CellText(cellValue)
    End If
End Function

' カンマ区切りの選択肢のどれかと同じか。
Private Function IsOneOf(ByVal target As String, ByVal choices As String) As Boolean
    Dim choice As Variant

    For Each choice In Split(choices, ",")
        If target = choice Then
            IsOneOf = True
            Exit Function
        End If
    Next choice
End Function
