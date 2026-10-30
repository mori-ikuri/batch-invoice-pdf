Attribute VB_Name = "modHistory"
'==============================================================================
' modHistory : 「実行履歴」への記録
'------------------------------------------------------------------------------
' ・ボタンを押すたびに1行足します(日時・操作・結果・作成件数・保存先・内容)。
' ・止まったときは「内容」に、どの段階・どの請求番号で止まったかを書きます。
' ・古い行は消しません。多くなったら、手で行を消すか別のブックへ移してください。
'==============================================================================
Option Explicit

Public Sub AppendHistory(ByVal operation As String, ByVal outcome As String, ByVal createdCount As Long, _
                         ByVal folderPath As String, ByVal detail As String)
    Dim ws As Worksheet
    Dim r As Long

    Set ws = shHistory
    ' フィルターで隠れた行があっても上書きしないよう、隠れた行も数えて最後の行を探す
    r = LastValueRow(ws, 1) + 1
    If r < HISTORY_FIRST_ROW Then r = HISTORY_FIRST_ROW

    ws.Cells(r, 1).Value = Now
    ws.Cells(r, 2).Value = operation
    ws.Cells(r, 3).Value = outcome
    ws.Cells(r, 4).Value = createdCount
    ws.Cells(r, 5).Value = folderPath
    ws.Cells(r, 6).Value = detail
End Sub
