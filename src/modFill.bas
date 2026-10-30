Attribute VB_Name = "modFill"
'==============================================================================
' modFill : ひな形シートへの流し込み
'------------------------------------------------------------------------------
' ・どのセルに何を入れるかは、ひな形シートの「名前の定義」で決めています
'   (数式タブ → 名前の管理。「差込_請求番号」などの名前で、範囲はひな形シート)。
'   セルを動かしても名前がついてくるので、レイアウトを変えてもコードは直さなくてよいです。
' ・明細欄の行数は「差込_品名」の行数で決まります。行を増やすときは、明細の4つの名前
'   (差込_品名・差込_数量・差込_単価・差込_税率)を同じ行数に広げてください。
' ・金額・小計・消費税・合計は、ひな形の式で計算します。マクロは値を入れるだけです。
'==============================================================================
Option Explicit

' シート名からシートを探します。なければ Nothing。
Public Function FindSheet(ByVal sheetTitle As String) As Worksheet
    Dim ws As Worksheet

    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = sheetTitle Then
            Set FindSheet = ws
            Exit Function
        End If
    Next ws
End Function

' ひな形の差し込み先の名前を、そのシートの範囲として返します。なければ Nothing。
Public Function TemplateRange(ByVal ws As Worksheet, ByVal rangeName As String) As Range
    On Error Resume Next   ' 名前がないときはエラーになるので、Nothing のまま返す
    Set TemplateRange = ws.Range(rangeName)
End Function

' ひな形に名前がそろっているかを調べ、明細欄に入る行数を返します(問題があれば 0)。
Public Function CheckTemplate(ByVal ws As Worksheet, ByVal errorList As Collection) As Long
    Dim nm As Variant
    Dim itemRows As Long
    Dim hasProblem As Boolean

    For Each nm In TemplateNames()
        If TemplateRange(ws, CStr(nm)) Is Nothing Then
            AddError errorList, ws.Name, 0, "", CStr(nm), _
                "ひな形に名前「" & nm & "」がありません。数式タブの「名前の管理」で、このシートのセルにつけてください。"
            hasProblem = True
        End If
    Next nm
    If hasProblem Then Exit Function

    itemRows = TemplateRange(ws, NAME_ITEM).Rows.Count
    For Each nm In Array(NAME_QUANTITY, NAME_UNIT_PRICE, NAME_TAX_RATE)
        If TemplateRange(ws, CStr(nm)).Rows.Count <> itemRows Then
            AddError errorList, ws.Name, 0, "", CStr(nm), _
                "明細の名前の行数がそろっていません(「" & NAME_ITEM & "」は " & itemRows & " 行、「" & nm & "」は " & _
                TemplateRange(ws, CStr(nm)).Rows.Count & " 行)。同じ行数にしてください。"
            hasProblem = True
        End If
    Next nm
    If Not hasProblem Then CheckTemplate = itemRows
End Function

' 請求1件をひな形に入れます(入力チェックが済んだものだけを渡します)。
Public Sub FillInvoice(ByVal ws As Worksheet, ByVal inv As clsInvoice)
    Dim d As clsDetail
    Dim i As Long

    ClearTemplate ws

    TemplateRange(ws, NAME_INVOICE_NO).Value = inv.InvoiceNo
    TemplateRange(ws, NAME_INVOICE_DATE).Value = CDate(inv.InvoiceDate)
    TemplateRange(ws, NAME_CUSTOMER).Value = inv.CustomerName
    TemplateRange(ws, NAME_HONORIFIC).Value = inv.Honorific
    TemplateRange(ws, NAME_DUE_DATE).Value = CDate(inv.DueDate)

    For Each d In inv.Details
        i = i + 1
        TemplateRange(ws, NAME_ITEM).Cells(i, 1).Value = d.ItemName
        TemplateRange(ws, NAME_QUANTITY).Cells(i, 1).Value = d.Quantity
        TemplateRange(ws, NAME_UNIT_PRICE).Cells(i, 1).Value = d.UnitPrice
        TemplateRange(ws, NAME_TAX_RATE).Cells(i, 1).Value = d.TaxRate
    Next d

    ' 計算方法が「手動」になっていても、金額・小計・消費税・合計を最新にする
    ws.Calculate
End Sub

' ひな形の差し込み先を空にします(式や書式はそのまま)。
Public Sub ClearTemplate(ByVal ws As Worksheet)
    Dim nm As Variant
    Dim rng As Range
    Dim c As Range

    For Each nm In TemplateNames()
        Set rng = TemplateRange(ws, CStr(nm))
        If Not rng Is Nothing Then
            For Each c In rng.Cells
                ' 結合したセルは、結合の範囲ごと消さないとエラーになる
                If c.MergeCells Then c.MergeArea.ClearContents Else c.ClearContents
            Next c
        End If
    Next nm
End Sub

' ひな形にそろっているべき名前の一覧
Private Function TemplateNames() As Variant
    TemplateNames = Array(NAME_INVOICE_NO, NAME_INVOICE_DATE, NAME_CUSTOMER, NAME_HONORIFIC, NAME_DUE_DATE, _
                          NAME_ITEM, NAME_QUANTITY, NAME_UNIT_PRICE, NAME_TAX_RATE)
End Function
