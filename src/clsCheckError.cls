VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCheckError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================
' clsCheckError : 入力チェックで見つけたエラー1件(「エラー一覧」の1行になる)
'==============================================================================
Option Explicit

Public SheetName As String      ' どのシートか
Public RowNo As Long            ' 何行目か(0 のときは行を出さない)
Public InvoiceNo As String      ' どの請求か(わかるとき)
Public ItemName As String       ' どの項目か
Public Message As String        ' 何が問題で、どう直すか
