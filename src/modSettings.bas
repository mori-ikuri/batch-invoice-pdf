Attribute VB_Name = "modSettings"
'==============================================================================
' modSettings : 「設定」シートの読み込み
'------------------------------------------------------------------------------
' ・A列の項目名で行を探し、B列の値を読みます。行の順番や位置は変えてもかまいません。
' ・設定を1つ足すときは、次の4か所を直します。
'     1. 「設定」シートに行を足す(A列:項目名、B列:値、C列:説明)
'     2. modConstants に項目名の定数(SET_...)を足す
'     3. 下の AppSettings に欄を足す
'     4. LoadSettings で読み込む(値が空ならエラーにするかもここで決める)
' ・値が正しいかどうか(選べる値か、フォルダがあるか など)は modValidate で調べます。
'==============================================================================
Option Explicit

Public Type AppSettings
    TemplateSheetName As String   ' PDFにするひな形シートの名前
    OutputFolderText As String    ' 保存先フォルダ(書かれたまま)
    OutputFolder As String        ' 保存先フォルダ(フルパス。決められないときは "")
    FileNameRule As String        ' 例: {請求番号}_{顧客名}.pdf
    IfExists As String            ' 同じ名前のファイルがあるとき:「止める」か「上書き」
    SaveAfterRun As String        ' 実行後にブックを保存:「する」か「しない」
    CompanyName As String         ' 請求書に載せる自社名
    RegistrationNo As String      ' 適格請求書発行事業者の登録番号
    TaxRounding As String         ' 消費税の端数処理
End Type

' 「設定」シートを読みます。行がない項目・値が空の項目は errorList に足します。
Public Function LoadSettings(ByVal errorList As Collection) As AppSettings
    Dim st As AppSettings

    st.TemplateSheetName = ReadSetting(SET_TEMPLATE_SHEET, errorList)
    st.OutputFolderText = ReadSetting(SET_OUTPUT_FOLDER, errorList)
    st.OutputFolder = ResolveFolder(st.OutputFolderText)
    st.FileNameRule = ReadSetting(SET_FILE_NAME_RULE, errorList)
    st.IfExists = ReadSetting(SET_IF_EXISTS, errorList)
    st.SaveAfterRun = ReadSetting(SET_SAVE_AFTER_RUN, errorList)
    st.CompanyName = ReadSetting(SET_COMPANY_NAME, errorList)
    st.RegistrationNo = ReadSetting(SET_REGISTRATION_NO, errorList)
    st.TaxRounding = ReadSetting(SET_TAX_ROUNDING, errorList)

    LoadSettings = st
End Function

' 項目名がある行の番号を返します。見つからなければ 0。
Public Function SettingRow(ByVal itemLabel As String) As Long
    Dim lastRow As Long
    Dim r As Long

    lastRow = shSettings.Cells(shSettings.Rows.Count, SETTINGS_LABEL_COLUMN).End(xlUp).Row
    For r = 1 To lastRow
        If Trim$(CellText(shSettings.Cells(r, SETTINGS_LABEL_COLUMN).Value)) = itemLabel Then
            SettingRow = r
            Exit Function
        End If
    Next r
End Function

' 保存先フォルダの書き方を、フルパスにして返します。
'   「C:\...」や「\\サーバー名\...」で始まるものはそのまま。それ以外は、このブックのある場所からの相対パス。
'   ブックが未保存のとき、OneDrive などの URL の上にあるときは、相対パスを決められないので "" を返します。
Public Function ResolveFolder(ByVal folderText As String) As String
    Dim folderPath As String

    folderPath = Trim$(folderText)
    If folderPath = "" Then Exit Function

    If folderPath Like "[A-Za-z]:\*" Or Left$(folderPath, 2) = "\\" Then
        ' フルパス
    ElseIf ThisWorkbook.Path = "" Or InStr(1, ThisWorkbook.Path, "://") > 0 Then
        Exit Function
    Else
        folderPath = ThisWorkbook.Path & "\" & folderPath
    End If

    ' 最後の「\」は取る(「C:\」のようなドライブだけのときは残す)
    Do While Len(folderPath) > 3 And Right$(folderPath, 1) = "\"
        folderPath = Left$(folderPath, Len(folderPath) - 1)
    Loop
    ResolveFolder = folderPath
End Function

' 項目の値を文字で返します。行がない・値が空のときは errorList に足します。
Private Function ReadSetting(ByVal itemLabel As String, ByVal errorList As Collection) As String
    Dim r As Long

    r = SettingRow(itemLabel)
    If r = 0 Then
        AddError errorList, shSettings.Name, 0, "", itemLabel, _
            "「設定」シートのA列に「" & itemLabel & "」の行がありません。項目名を書き換えていないか確かめてください。"
        Exit Function
    End If

    ReadSetting = Trim$(CellText(shSettings.Cells(r, SETTINGS_VALUE_COLUMN).Value))
    If ReadSetting = "" Then
        AddError errorList, shSettings.Name, r, "", itemLabel, "値が入っていません。B列に入れてください。"
    End If
End Function
