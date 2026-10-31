VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================
' clsDetail : 明細1行分(「明細」シートの1行)
'------------------------------------------------------------------------------
' 日付・数量・単価・税率はセルの値のまま持ちます。日付・数値として正しいかは modValidate で調べ、
' 税率は正しければ 0.1 か 0.08 にそろえます(ひな形の式が税率ごとに合計するため)。
'==============================================================================
Option Explicit

Public InvoiceNo As String      ' 請求番号(請求一覧と結びつける)
Public TradeDate As Variant     ' 日付(取引年月日。請求書には 月/日 で出す)
Public ItemName As String       ' 品名
Public Quantity As Variant      ' 数量
Public UnitPrice As Variant     ' 単価(値引きは負の数でもよい)
Public TaxRate As Variant       ' 税率(10% か 8%)
Public RowNo As Long            ' 「明細」の行番号(エラー一覧に出す)
