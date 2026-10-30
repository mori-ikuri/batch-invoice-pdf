Attribute VB_Name = "modData"
'==============================================================================
' modData : 「請求一覧」「明細」の読み込みと、セルの値の扱い
'------------------------------------------------------------------------------
' ・1行目の見出しの文字で列を探します。列の順番を入れ替えても動きます。
' ・すべての列が空の行は飛ばします(途中に空の行があってもかまいません)。
' ・ここでは値をそのまま読むだけです。正しいかどうかは modValidate で調べます。
'==============================================================================
Option Explicit

' 「請求一覧」を読み、clsInvoice の集まりを返します。
' 見出しが足りないときは errorList に足し、空の集まりを返します。
Public Function ReadInvoices(ByVal errorList As Collection) As Collection
    Dim ws As Worksheet
    Dim invoices As Collection
    Dim inv As clsInvoice
    Dim cNo As Long, cInvoiceDate As Long, cCustomer As Long, cHonorific As Long, cDue As Long
    Dim lastRow As Long
    Dim r As Long

    Set ws = shInvoices
    Set invoices = New Collection
    Set ReadInvoices = invoices

    cNo = FindColumn(ws, COL_INVOICE_NO, errorList)
    cInvoiceDate = FindColumn(ws, COL_INVOICE_DATE, errorList)
    cCustomer = FindColumn(ws, COL_CUSTOMER, errorList)
    cHonorific = FindColumn(ws, COL_HONORIFIC, errorList)
    cDue = FindColumn(ws, COL_DUE_DATE, errorList)
    If cNo = 0 Or cInvoiceDate = 0 Or cCustomer = 0 Or cHonorific = 0 Or cDue = 0 Then Exit Function

    lastRow = LastDataRow(ws, Array(cNo, cInvoiceDate, cCustomer, cHonorific, cDue))
    For r = HEADER_ROW + 1 To lastRow
        If Not IsBlankRow(ws, r, Array(cNo, cInvoiceDate, cCustomer, cHonorific, cDue)) Then
            Set inv = New clsInvoice
            inv.RowNo = r
            inv.InvoiceNo = Trim$(CellText(ws.Cells(r, cNo).Value))
            inv.InvoiceDate = ws.Cells(r, cInvoiceDate).Value
            inv.CustomerName = Trim$(CellText(ws.Cells(r, cCustomer).Value))
            inv.Honorific = Trim$(CellText(ws.Cells(r, cHonorific).Value))
            inv.DueDate = ws.Cells(r, cDue).Value
            invoices.Add inv
        End If
    Next r
End Function

' 「明細」を読み、clsDetail の集まりを返します。
' 見出しが足りないときは errorList に足し、空の集まりを返します。
Public Function ReadDetails(ByVal errorList As Collection) As Collection
    Dim ws As Worksheet
    Dim allDetails As Collection
    Dim d As clsDetail
    Dim cNo As Long, cItem As Long, cQuantity As Long, cPrice As Long, cRate As Long
    Dim lastRow As Long
    Dim r As Long

    Set ws = shDetails
    Set allDetails = New Collection
    Set ReadDetails = allDetails

    cNo = FindColumn(ws, COL_INVOICE_NO, errorList)
    cItem = FindColumn(ws, COL_ITEM, errorList)
    cQuantity = FindColumn(ws, COL_QUANTITY, errorList)
    cPrice = FindColumn(ws, COL_UNIT_PRICE, errorList)
    cRate = FindColumn(ws, COL_TAX_RATE, errorList)
    If cNo = 0 Or cItem = 0 Or cQuantity = 0 Or cPrice = 0 Or cRate = 0 Then Exit Function

    lastRow = LastDataRow(ws, Array(cNo, cItem, cQuantity, cPrice, cRate))
    For r = HEADER_ROW + 1 To lastRow
        If Not IsBlankRow(ws, r, Array(cNo, cItem, cQuantity, cPrice, cRate)) Then
            Set d = New clsDetail
            d.RowNo = r
            d.InvoiceNo = Trim$(CellText(ws.Cells(r, cNo).Value))
            d.ItemName = Trim$(CellText(ws.Cells(r, cItem).Value))
            d.Quantity = ws.Cells(r, cQuantity).Value
            d.UnitPrice = ws.Cells(r, cPrice).Value
            d.TaxRate = ws.Cells(r, cRate).Value
            allDetails.Add d
        End If
    Next r
End Function

' 見出しの文字がある列の番号を返します。見つからなければ 0。
Public Function HeaderColumn(ByVal ws As Worksheet, ByVal header As String) As Long
    Dim lastCol As Long
    Dim c As Long

    lastCol = ws.Cells(HEADER_ROW, ws.Columns.Count).End(xlToLeft).Column
    For c = 1 To lastCol
        If Trim$(CellText(ws.Cells(HEADER_ROW, c).Value)) = header Then
            HeaderColumn = c
            Exit Function
        End If
    Next c
End Function

' セルの値を文字にします。空なら ""、#N/A などのエラー値なら "#エラー" を返します。
Public Function CellText(ByVal cellValue As Variant) As String
    If IsError(cellValue) Then
        CellText = "#エラー"
    ElseIf IsEmpty(cellValue) Or IsNull(cellValue) Then
        CellText = ""
    Else
        CellText = CStr(cellValue)
    End If
End Function

' セルの値が数値か(文字として入った数字は数値と見なしません)。
Public Function IsNumberValue(ByVal cellValue As Variant) As Boolean
    Select Case VarType(cellValue)
        Case vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal
            IsNumberValue = True
    End Select
End Function

' セルの値が日付として読めるか(日付の値か、日付として読める文字)。
Public Function IsDateValue(ByVal cellValue As Variant) As Boolean
    Select Case VarType(cellValue)
        Case vbDate
            IsDateValue = True
        Case vbString
            IsDateValue = IsDate(cellValue)
    End Select
End Function

' 見出しの列を探します。見つからなければ errorList に足して 0 を返します。
Private Function FindColumn(ByVal ws As Worksheet, ByVal header As String, ByVal errorList As Collection) As Long
    FindColumn = HeaderColumn(ws, header)
    If FindColumn = 0 Then
        AddError errorList, ws.Name, HEADER_ROW, "", header, _
            "1行目に見出し「" & header & "」がありません。見出しの文字を確かめてください。"
    End If
End Function

' 指定した列のうち、いちばん下まで値がある行の番号を返します。
Private Function LastDataRow(ByVal ws As Worksheet, ByVal cols As Variant) As Long
    Dim c As Variant
    Dim r As Long

    LastDataRow = HEADER_ROW
    For Each c In cols
        r = ws.Cells(ws.Rows.Count, CLng(c)).End(xlUp).Row
        If r > LastDataRow Then LastDataRow = r
    Next c
End Function

' 指定した列がすべて空の行か。
Private Function IsBlankRow(ByVal ws As Worksheet, ByVal r As Long, ByVal cols As Variant) As Boolean
    Dim c As Variant

    For Each c In cols
        If Trim$(CellText(ws.Cells(r, CLng(c)).Value)) <> "" Then Exit Function
    Next c
    IsBlankRow = True
End Function
