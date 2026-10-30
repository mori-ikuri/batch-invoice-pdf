VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsInvoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================
' clsInvoice : 請求1件分(「請求一覧」の1行と、それにつながる明細)
'------------------------------------------------------------------------------
' 日付はセルの値のまま持ちます。日付として正しいかは modValidate で調べます。
'==============================================================================
Option Explicit

Public InvoiceNo As String      ' 請求番号(前後の空白は取ったもの)
Public InvoiceDate As Variant   ' 請求日
Public CustomerName As String   ' 顧客名
Public Honorific As String      ' 敬称(御中・様 など)
Public DueDate As Variant       ' 支払期限
Public RowNo As Long            ' 「請求一覧」の行番号(エラー一覧に出す)
Public Details As Collection    ' この請求の明細(clsDetail)
Public FileName As String       ' PDFのファイル名(入力チェックのときに決める)

Private Sub Class_Initialize()
    Set Details = New Collection
End Sub
