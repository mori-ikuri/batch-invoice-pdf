Attribute VB_Name = "modPdf"
'==============================================================================
' modPdf : ファイル名を決めて、PDFに保存する
'------------------------------------------------------------------------------
' ・ファイル名は、「設定」のファイル名ルールの差し込み({請求番号} など)を置き換えて作ります。
'   差し込みを足すときは、modConstants に TAG_... を足し、BuildFileName と UnknownTag に足します。
' ・PDFは Excel の「PDFとして保存」と同じ機能(ExportAsFixedFormat)で、このPCのフォルダにだけ
'   保存します。ネット通信・外部サービス・アドインは使いません。
' ・保存先フォルダがないときに作る手続き(CreateFolder)もここにあります。作るのは最後の1階層だけです。
'==============================================================================
Option Explicit

' ファイル名ルールに請求の値を差し込みます。
Public Function BuildFileName(ByVal rule As String, ByVal inv As clsInvoice) As String
    Dim result As String

    result = Replace(rule, TAG_INVOICE_NO, inv.InvoiceNo)
    result = Replace(result, TAG_CUSTOMER, inv.CustomerName)
    If IsDateValue(inv.InvoiceDate) Then
        result = Replace(result, TAG_INVOICE_DATE, Format$(CDate(inv.InvoiceDate), FILE_DATE_FORMAT))
    Else
        result = Replace(result, TAG_INVOICE_DATE, "")
    End If
    BuildFileName = result
End Function

' ルールの中の { } で囲んだ差し込みのうち、使えないものを返します(なければ "")。
Public Function UnknownTag(ByVal rule As String) As String
    Dim startPos As Long
    Dim endPos As Long
    Dim tag As String

    startPos = InStr(1, rule, "{")
    Do While startPos > 0
        endPos = InStr(startPos + 1, rule, "}")
        If endPos = 0 Then
            UnknownTag = Mid$(rule, startPos)
            Exit Function
        End If
        tag = Mid$(rule, startPos, endPos - startPos + 1)
        If tag <> TAG_INVOICE_NO And tag <> TAG_CUSTOMER And tag <> TAG_INVOICE_DATE Then
            UnknownTag = tag
            Exit Function
        End If
        startPos = InStr(endPos + 1, rule, "{")
    Loop
End Function

' ファイル名に使えない文字(\ / : * ? " < > | と改行などの制御文字)があるか。
Public Function HasInvalidFileChar(ByVal fileTitle As String) As Boolean
    Dim i As Long
    Dim ch As String

    For i = 1 To Len(fileTitle)
        ch = Mid$(fileTitle, i, 1)
        ' AscW は漢字などで負の数を返すので、&HFFFF& で 0 から 65535 の数にそろえてから比べる
        If InStr(1, INVALID_FILE_CHARS, ch) > 0 Or (AscW(ch) And &HFFFF&) < 32 Then
            HasInvalidFileChar = True
            Exit Function
        End If
    Next i
End Function

Public Function FolderExists(ByVal folderPath As String) As Boolean
    On Error Resume Next   ' ないときは GetAttr がエラーになり、False のまま返る
    FolderExists = ((GetAttr(folderPath) And vbDirectory) = vbDirectory)
End Function

Public Function FileExists(ByVal filePath As String) As Boolean
    On Error Resume Next   ' ないときは GetAttr がエラーになり、False のまま返る
    FileExists = ((GetAttr(filePath) And vbDirectory) = 0)
End Function

' フォルダがなく、その1つ上のフォルダはあるか(最後の1階層だけ作れば使えるか)。
' 1つ上もないとき(ドライブ名の打ち間違いなど)は False。打ち間違いのまま、思わぬ場所に作らないためです。
Public Function CanCreateFolder(ByVal folderPath As String) As Boolean
    Dim parentPath As String

    If folderPath = "" Or FolderExists(folderPath) Then Exit Function
    parentPath = ParentFolder(folderPath)
    If parentPath <> "" Then CanCreateFolder = FolderExists(parentPath)
End Function

' フォルダの最後の1階層だけを作ります(画面は出しません)。途中のフォルダまでは作りません。
' 作れたとき・もうあるときは ""、作れなかったときは理由を返します。
Public Function CreateFolder(ByVal folderPath As String) As String
    If FolderExists(folderPath) Then Exit Function
    If Not CanCreateFolder(folderPath) Then
        CreateFolder = "1つ上のフォルダがないため、作りませんでした。"
        Exit Function
    End If

    On Error Resume Next   ' 作れないとき(書き込みの権限がない など)は、止まらずに理由を返す
    MkDir folderPath
    If Err.Number <> 0 Then CreateFolder = Err.Description & "(エラー番号 " & Err.Number & ")"
End Function

' 1つ上のフォルダを返します(「C:\請求書PDF」なら「C:\」)。上がないとき(「\\サーバー\共有」など)は ""。
Private Function ParentFolder(ByVal folderPath As String) As String
    Dim pos As Long

    pos = InStrRev(folderPath, "\")
    If pos = 0 Then Exit Function
    ParentFolder = Left$(folderPath, pos - 1)
    If ParentFolder Like "[A-Za-z]:" Then
        ParentFolder = ParentFolder & "\"   ' ドライブだけのときは「C:\」の形にする
    ElseIf Left$(folderPath, 2) = "\\" And UBound(Split(ParentFolder, "\")) < 3 Then
        ParentFolder = ""                   ' 「\\サーバー」だけでは、フォルダではない
    End If
End Function

' シートをPDFに保存します。失敗したときはエラーを起こし、呼んだ側(ExecuteBatch)で受け止めます。
Public Sub ExportPdf(ByVal ws As Worksheet, ByVal filePath As String, ByVal overwrite As Boolean)
    If FileExists(filePath) Then
        If Not overwrite Then
            Err.Raise vbObjectError + 1001, "ExportPdf", "同じ名前のファイルがあります:" & filePath
        End If
        ' 上書きのときは先に消す。PDFを開いたままだと、ここで止まります(書き込みできません)。
        Kill filePath
    End If

    ' IncludeDocProperties:=False … ブックの作成者名などをPDFに入れない
    ws.ExportAsFixedFormat Type:=xlTypePDF, Filename:=filePath, Quality:=xlQualityStandard, _
                           IncludeDocProperties:=False, IgnorePrintAreas:=False, OpenAfterPublish:=False

    If Not FileExists(filePath) Then
        Err.Raise vbObjectError + 1002, "ExportPdf", "PDFができませんでした:" & filePath
    End If
End Sub
