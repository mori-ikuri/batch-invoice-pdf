Attribute VB_Name = "modMain"
'==============================================================================
' modMain : ボタンから呼ぶマクロと、全体の流れ
'------------------------------------------------------------------------------
' ボタンとマクロ(「使い方」シート)
'   入力をチェック        → RunCheckOnly
'   PDFを作成             → RunCreatePdfs
'   1件だけひな形に表示   → RunPreview
'
' 流れ(ExecuteBatch)
'   設定の読み込み(modSettings) → データの読み込み(modData)
'   → 入力チェック(modValidate)  ※エラーが1件でもあれば、ここで止まり1件も作らない
'   → 1件ずつ ひな形へ流し込み(modFill) → PDF保存(modPdf)
'   → 実行履歴へ記録(modHistory)
' 保存先フォルダがないときは、その前に Run... が「このフォルダを作りますか?」と聞きます
' (1つ上のフォルダがあるときだけ。作るのは最後の1階層だけ)。
'
' メッセージを出すのは Run... の3つだけです。ExecuteBatch・ExecutePreview・OutputFolderToCreate と、
' フォルダを作る modPdf.CreateFolder は画面を出さないので、テストの道具(tests\run-tests.ps1)からも呼べます。
'==============================================================================
Option Explicit

' 直前の実行の結果。メッセージの出し分けに使います。
Public LastResult As String          ' RESULT_OK など
Public LastCreatedCount As Long      ' 作成したPDFの件数
Public LastOutputFolder As String    ' 保存先フォルダ
Public LastTemplateSheet As String   ' ひな形に表示したシート(RunPreview 用)

'------------------------------------------------------------------------------
' ボタンから呼ぶマクロ
'------------------------------------------------------------------------------

' 「入力をチェック」:PDFは作らず、入力チェックだけをします。
Public Sub RunCheckOnly()
    OfferToCreateFolder
    ShowResult ExecuteBatch(False)
End Sub

' 「PDFを作成」:入力チェックのあと、確認してからPDFを作ります。
Public Sub RunCreatePdfs()
    OfferToCreateFolder
    ShowResult ExecuteBatch(True, True)
End Sub

' 「1件だけひな形に表示」:指定した請求をひな形に入れて見せます(PDFは作りません)。
Public Sub RunPreview()
    Dim defaultNo As String
    Dim invoiceNumber As String
    Dim message As String

    ' 最初に入れておく番号:「請求一覧」「明細」で選んでいる行(Alt+F8 から実行したとき)、
    ' なければ「請求一覧」の先頭の請求番号(「使い方」シートのボタンから押したとき)
    defaultNo = SelectedInvoiceNo()
    If defaultNo = "" Then defaultNo = FirstInvoiceNo()

    invoiceNumber = Trim$(InputBox("ひな形に表示する請求番号を入れてください。(PDFは作りません)", _
                                   TOOL_TITLE, defaultNo))
    If invoiceNumber = "" Then Exit Sub

    message = ExecutePreview(invoiceNumber)
    If LastResult = RESULT_OK Then
        ThisWorkbook.Worksheets(LastTemplateSheet).Activate
    Else
        ShowResult message
    End If
End Sub

'------------------------------------------------------------------------------
' 本体
'------------------------------------------------------------------------------

' 入力チェックをし、createPdfs が True ならPDFを作ります。結果の文章を返します。
' askBeforeCreate が True のときだけ、作る前に確認のメッセージを出します。
Public Function ExecuteBatch(ByVal createPdfs As Boolean, Optional ByVal askBeforeCreate As Boolean = False) As String
    Dim st As AppSettings
    Dim errorList As Collection
    Dim invoices As Collection
    Dim allDetails As Collection
    Dim inv As clsInvoice
    Dim templateWs As Worksheet
    Dim operation As String
    Dim stage As String           ' いまの段階(止まったときに出す)
    Dim currentNo As String       ' いま作っている請求の番号(止まったときに出す)
    Dim createdCount As Long
    Dim message As String
    Dim errNumber As Long
    Dim errText As String
    Dim answer As VbMsgBoxResult

    operation = IIf(createPdfs, OP_CREATE, OP_CHECK)
    LastResult = ""
    LastCreatedCount = 0
    LastOutputFolder = ""

    On Error GoTo Failed
    BeginWork

    stage = "設定の読み込み"
    Set errorList = New Collection
    st = LoadSettings(errorList)
    LastOutputFolder = st.OutputFolder

    stage = "データの読み込み"
    Set invoices = ReadInvoices(errorList)
    Set allDetails = ReadDetails(errorList)

    stage = "入力チェック"
    ValidateAll st, invoices, allDetails, errorList
    WriteErrorList errorList, invoices.Count

    If errorList.Count > 0 Then
        LastResult = RESULT_INVALID
        message = "入力にエラーが " & errorList.Count & " 件あります。" & vbCrLf & _
                  "「エラー一覧」シートを見て、上から順に直してください。"
        If createPdfs Then message = message & vbCrLf & "PDFは1件も作成していません。"
        AppendHistory operation, RESULT_INVALID, 0, st.OutputFolder, "エラー " & errorList.Count & " 件"
        GoTo Finish
    End If

    If Not createPdfs Then
        LastResult = RESULT_OK
        message = "エラーはありません。請求 " & invoices.Count & " 件のPDFを作成できます。"
        AppendHistory operation, RESULT_OK, 0, st.OutputFolder, "請求 " & invoices.Count & " 件、エラーなし"
        GoTo Finish
    End If

    If askBeforeCreate Then
        EndWork   ' 確認のあいだは、画面の書き換えとカーソル(砂時計)を元に戻す
        answer = MsgBox(invoices.Count & " 件の請求書PDFを作成します。" & vbCrLf & _
                        "保存先:" & st.OutputFolder & vbCrLf & _
                        IIf(st.IfExists = IF_EXISTS_OVERWRITE, "(同じ名前のファイルは置き換えます)" & vbCrLf, "") & _
                        vbCrLf & "よろしいですか?", vbOKCancel + vbQuestion, TOOL_TITLE)
        If answer <> vbOK Then
            LastResult = RESULT_CANCELLED
            message = "PDFの作成を取りやめました。"
            AppendHistory operation, RESULT_CANCELLED, 0, st.OutputFolder, ""
            GoTo Finish
        End If
        BeginWork
    End If

    stage = "PDF作成"
    Set templateWs = FindSheet(st.TemplateSheetName)
    For Each inv In invoices
        currentNo = inv.InvoiceNo
        Application.StatusBar = "請求書PDFを作成中 " & (createdCount + 1) & "/" & invoices.Count & "(" & currentNo & ")"
        FillInvoice templateWs, inv
        ExportPdf templateWs, st.OutputFolder & "\" & inv.FileName, (st.IfExists = IF_EXISTS_OVERWRITE)
        createdCount = createdCount + 1
    Next inv
    currentNo = ""
    ClearTemplate templateWs

    LastResult = RESULT_OK
    LastCreatedCount = createdCount
    message = createdCount & " 件の請求書PDFを作成しました。" & vbCrLf & "保存先:" & st.OutputFolder
    AppendHistory operation, RESULT_OK, createdCount, st.OutputFolder, ""

Finish:
    message = message & SaveWorkbookIfSet(st)
    EndWork
    ExecuteBatch = message
    Exit Function

Failed:
    ' 想定外のエラー:どの段階・どの請求番号で止まったかを、画面と実行履歴に残す
    errNumber = Err.Number
    errText = Err.Description
    Resume Aborted

Aborted:
    On Error Resume Next   ' 後始末の途中で、もう一度止まらないようにする
    LastResult = RESULT_ABORTED
    LastCreatedCount = createdCount
    message = "想定外のエラーで止まりました。" & vbCrLf & _
              "段階:" & stage & vbCrLf
    If currentNo <> "" Then
        message = message & "請求番号:" & currentNo & "(ひな形シートに、この請求を表示したままにしています)" & vbCrLf
    End If
    message = message & "内容:" & errText & "(エラー番号 " & errNumber & ")"
    If createPdfs Then message = message & vbCrLf & "それまでに作成したPDF:" & createdCount & " 件"
    AppendHistory operation, RESULT_ABORTED, createdCount, st.OutputFolder, _
        "段階:" & stage & IIf(currentNo <> "", " / 請求番号:" & currentNo, "") & _
        " / " & errText & "(エラー番号 " & errNumber & ")"
    GoTo Finish
End Function

' 指定した請求番号の請求を、入力チェックをしてからひな形に入れます。結果の文章を返します。
' ほかの請求のエラーや、保存先の設定は見ません(見た目を確かめるためのものなので)。
Public Function ExecutePreview(ByVal invoiceNumber As String) As String
    Dim st As AppSettings
    Dim ignored As Collection
    Dim errorList As Collection
    Dim invoices As Collection
    Dim allDetails As Collection
    Dim inv As clsInvoice
    Dim target As clsInvoice
    Dim d As clsDetail
    Dim templateWs As Worksheet
    Dim capacity As Long

    LastResult = ""
    LastTemplateSheet = ""
    On Error GoTo Failed

    ' 表示に要る設定はひな形シート名だけなので、ほかの設定のエラーはここでは数えない
    Set ignored = New Collection
    st = LoadSettings(ignored)

    Set errorList = New Collection
    If st.TemplateSheetName = "" Then
        AddError errorList, shSettings.Name, SettingRow(SET_TEMPLATE_SHEET), "", SET_TEMPLATE_SHEET, "値が入っていません。"
    Else
        Set templateWs = FindSheet(st.TemplateSheetName)
        If templateWs Is Nothing Then
            AddError errorList, shSettings.Name, SettingRow(SET_TEMPLATE_SHEET), "", SET_TEMPLATE_SHEET, _
                "「" & st.TemplateSheetName & "」というシートがありません。"
        Else
            capacity = CheckTemplate(templateWs, errorList)
        End If
    End If

    Set invoices = ReadInvoices(errorList)
    Set allDetails = ReadDetails(errorList)
    For Each inv In invoices
        If StrComp(inv.InvoiceNo, invoiceNumber, vbTextCompare) = 0 Then
            Set target = inv
            Exit For
        End If
    Next inv

    If target Is Nothing Then
        AddError errorList, shInvoices.Name, 0, invoiceNumber, COL_INVOICE_NO, "「請求一覧」にこの請求番号がありません。"
    Else
        CheckInvoice target, errorList
        For Each d In allDetails
            If StrComp(d.InvoiceNo, invoiceNumber, vbTextCompare) = 0 Then
                CheckDetail d, errorList
                target.Details.Add d
            End If
        Next d
        If target.Details.Count = 0 Then
            AddError errorList, shInvoices.Name, target.RowNo, target.InvoiceNo, "明細", "明細が1行もありません。"
        ElseIf capacity > 0 And target.Details.Count > capacity Then
            AddError errorList, shInvoices.Name, target.RowNo, target.InvoiceNo, "明細", _
                "明細が " & target.Details.Count & " 行あり、ひな形の明細欄(" & capacity & " 行)に入りきりません。"
        End If
    End If

    If errorList.Count > 0 Then
        WriteErrorList errorList, invoices.Count
        LastResult = RESULT_INVALID
        ExecutePreview = "請求番号 " & invoiceNumber & " をひな形に表示できません。" & vbCrLf & _
                         "エラーが " & errorList.Count & " 件あります。「エラー一覧」シートを見てください。"
        Exit Function
    End If

    FillInvoice templateWs, target
    LastResult = RESULT_OK
    LastTemplateSheet = templateWs.Name
    ExecutePreview = "請求番号 " & invoiceNumber & " を「" & templateWs.Name & "」シートに表示しました。"
    Exit Function

Failed:
    LastResult = RESULT_ABORTED
    ExecutePreview = "想定外のエラーで止まりました。" & vbCrLf & _
                     "請求番号:" & invoiceNumber & vbCrLf & _
                     "内容:" & Err.Description & "(エラー番号 " & Err.Number & ")"
End Function

' 作るかを聞く保存先フォルダのフルパスを返します。フォルダがなく、その1つ上のフォルダはあるときだけで、
' ほかは ""(もうある・1つ上もない・設定を読めない など。そのときは入力チェックがエラー一覧に出します)。
Public Function OutputFolderToCreate() As String
    Dim st As AppSettings

    On Error Resume Next   ' 設定を読めないときは "" のまま
    st = LoadSettings(New Collection)
    If CanCreateFolder(st.OutputFolder) Then OutputFolderToCreate = st.OutputFolder
End Function

'------------------------------------------------------------------------------
' 小さな道具
'------------------------------------------------------------------------------

' 保存先フォルダがなく、その1つ上のフォルダはあるときに、作るかを聞きます。[はい] なら作ります。
' [いいえ] のときや作れなかったときは、このあとの入力チェックが「エラー一覧」に出して止めます。
Private Sub OfferToCreateFolder()
    Dim folderPath As String
    Dim reason As String

    folderPath = OutputFolderToCreate()
    If folderPath = "" Then Exit Sub
    If MsgBox("保存先フォルダがありません:" & folderPath & "。" & vbCrLf & "このフォルダを作りますか?", _
              vbYesNo + vbQuestion, TOOL_TITLE) <> vbYes Then Exit Sub

    reason = CreateFolder(folderPath)
    If reason <> "" Then
        MsgBox "保存先フォルダを作れませんでした:" & folderPath & vbCrLf & "内容:" & reason & vbCrLf & vbCrLf & _
               "設定の保存先を、書き込めるフォルダに直してください。", vbExclamation, TOOL_TITLE
    End If
End Sub

' 結果に合わせてメッセージを出し、見るべきシートを開きます。
Private Sub ShowResult(ByVal message As String)
    Select Case LastResult
        Case RESULT_OK
            If LastCreatedCount > 0 Then
                If MsgBox(message & vbCrLf & vbCrLf & "保存先のフォルダを開きますか?", _
                          vbYesNo + vbInformation, TOOL_TITLE) = vbYes Then
                    Shell "explorer.exe """ & LastOutputFolder & """", vbNormalFocus
                End If
            Else
                MsgBox message, vbInformation, TOOL_TITLE
            End If
        Case RESULT_INVALID
            shErrors.Activate
            MsgBox message, vbExclamation, TOOL_TITLE
        Case RESULT_CANCELLED
            MsgBox message, vbInformation, TOOL_TITLE
        Case Else
            shHistory.Activate
            MsgBox message, vbCritical, TOOL_TITLE
    End Select
End Sub

' 「請求一覧」の先頭の請求番号を返します(読めないときは "")。
Private Function FirstInvoiceNo() As String
    Dim invoices As Collection

    On Error Resume Next   ' 読めないときは "" のまま
    Set invoices = ReadInvoices(New Collection)
    If invoices Is Nothing Then Exit Function
    If invoices.Count > 0 Then FirstInvoiceNo = invoices(1).InvoiceNo
End Function

' 「請求一覧」「明細」で選んでいる行の請求番号を返します(ほかのシートなら "")。
Private Function SelectedInvoiceNo() As String
    Dim ws As Object
    Dim c As Long

    On Error Resume Next   ' 選んでいるものによっては取れないので、そのときは "" のまま
    Set ws = ActiveSheet
    If ws Is shInvoices Or ws Is shDetails Then
        c = HeaderColumn(ws, COL_INVOICE_NO)
        If c > 0 And ActiveCell.Row > HEADER_ROW Then
            SelectedInvoiceNo = Trim$(CellText(ws.Cells(ActiveCell.Row, c).Value))
        End If
    End If
End Function

' 設定が「する」なら、実行履歴を残すためにブックを上書き保存します。
' 保存できなかったときは、メッセージに足す文章を返します。
Private Function SaveWorkbookIfSet(ByRef st As AppSettings) As String
    If st.SaveAfterRun <> YES_TEXT Then Exit Function
    If ThisWorkbook.ReadOnly Or ThisWorkbook.Path = "" Then
        SaveWorkbookIfSet = vbCrLf & "(ブックが読み取り専用か未保存のため、上書き保存はしていません)"
        Exit Function
    End If

    On Error Resume Next
    ThisWorkbook.Save
    If Err.Number <> 0 Then
        SaveWorkbookIfSet = vbCrLf & "(ブックを上書き保存できませんでした:" & Err.Description & ")"
    End If
End Function

' 作業中は画面の書き換えを止め、砂時計にします。
Private Sub BeginWork()
    Application.ScreenUpdating = False
    Application.Cursor = xlWait
End Sub

Private Sub EndWork()
    Application.StatusBar = False
    Application.Cursor = xlDefault
    Application.ScreenUpdating = True
End Sub
