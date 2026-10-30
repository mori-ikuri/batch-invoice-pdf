Attribute VB_Name = "modConstants"
'==============================================================================
' modConstants : 決め打ちの値をまとめる場所
'------------------------------------------------------------------------------
' ・シートの見出し、「設定」シートの項目名、選べる値、ひな形の名前などを集めています。
' ・シートの見出しや項目名を変えたときは、ここの文字も同じに直してください。
' ・使う人が変える値(保存先、自社情報など)は、ここではなく「設定」シートに置きます。
'==============================================================================
Option Explicit

' メッセージの題名
Public Const TOOL_TITLE As String = "請求書一括PDF作成"

'--- 「請求一覧」「明細」の見出し(1行目)
'    見出しの文字で列を探すので、列の順番は入れ替えてもかまいません。
Public Const HEADER_ROW As Long = 1
Public Const COL_INVOICE_NO As String = "請求番号"
Public Const COL_INVOICE_DATE As String = "請求日"
Public Const COL_CUSTOMER As String = "顧客名"
Public Const COL_HONORIFIC As String = "敬称"
Public Const COL_DUE_DATE As String = "支払期限"
Public Const COL_ITEM As String = "品名"
Public Const COL_QUANTITY As String = "数量"
Public Const COL_UNIT_PRICE As String = "単価"
Public Const COL_TAX_RATE As String = "税率"

'--- 「設定」シートの項目名(A列)と、値の列(B列)
Public Const SETTINGS_LABEL_COLUMN As Long = 1
Public Const SETTINGS_VALUE_COLUMN As Long = 2
Public Const SET_TEMPLATE_SHEET As String = "ひな形シート名"
Public Const SET_OUTPUT_FOLDER As String = "保存先フォルダ"
Public Const SET_FILE_NAME_RULE As String = "ファイル名ルール"
Public Const SET_IF_EXISTS As String = "同名ファイルがあるとき"
Public Const SET_SAVE_AFTER_RUN As String = "実行後にブックを保存"
Public Const SET_COMPANY_NAME As String = "自社名"
Public Const SET_REGISTRATION_NO As String = "登録番号"
Public Const SET_TAX_ROUNDING As String = "消費税の端数処理"

'--- 設定で選べる値
Public Const IF_EXISTS_STOP As String = "止める"
Public Const IF_EXISTS_OVERWRITE As String = "上書き"
Public Const YES_TEXT As String = "する"
Public Const NO_TEXT As String = "しない"
' 消費税の端数処理(実際の計算は、ひな形シートの式がこの文字を見て行います)
Public Const ROUNDING_CHOICES As String = "切り捨て,四捨五入,切り上げ"
' 登録番号の形(T のあとに数字13桁)
Public Const REGISTRATION_NO_PATTERN As String = "T#############"

'--- ファイル名ルールの差し込み
Public Const TAG_INVOICE_NO As String = "{請求番号}"
Public Const TAG_CUSTOMER As String = "{顧客名}"
Public Const TAG_INVOICE_DATE As String = "{請求日}"
Public Const FILE_DATE_FORMAT As String = "yyyymmdd"
' Windows のファイル名に使えない文字
Public Const INVALID_FILE_CHARS As String = "\/:*?""<>|"
' 保存先とファイル名を合わせた長さの上限(Windows の既定)
Public Const MAX_PATH_LENGTH As Long = 259

'--- 明細の税率(この2つだけを受け付けます)
Public Const TAX_RATE_STANDARD As Double = 0.1
Public Const TAX_RATE_REDUCED As Double = 0.08

'--- ひな形シートの「名前の定義」(差し込み先)
Public Const NAME_INVOICE_NO As String = "差込_請求番号"
Public Const NAME_INVOICE_DATE As String = "差込_請求日"
Public Const NAME_CUSTOMER As String = "差込_顧客名"
Public Const NAME_HONORIFIC As String = "差込_敬称"
Public Const NAME_DUE_DATE As String = "差込_支払期限"
Public Const NAME_ITEM As String = "差込_品名"
Public Const NAME_QUANTITY As String = "差込_数量"
Public Const NAME_UNIT_PRICE As String = "差込_単価"
Public Const NAME_TAX_RATE As String = "差込_税率"

'--- 「エラー一覧」「実行履歴」の書き始めの位置
Public Const ERROR_STATUS_CELL As String = "A2"
Public Const ERROR_FIRST_ROW As Long = 5
Public Const ERROR_COLUMNS As Long = 6
Public Const HISTORY_FIRST_ROW As Long = 4

'--- 実行履歴の「操作」と「結果」
Public Const OP_CHECK As String = "入力チェック"
Public Const OP_CREATE As String = "PDF作成"
Public Const RESULT_OK As String = "完了"
Public Const RESULT_INVALID As String = "エラーあり"
Public Const RESULT_CANCELLED As String = "取りやめ"
Public Const RESULT_ABORTED As String = "中断"
