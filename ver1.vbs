Option Explicit

'────────────────────────────────────────
'  シンプル処理：ドラッグ&ドロップでExcelファイルを受け取り、
'  A3とA4の値をポップアップで表示
'────────────────────────────────────────

Dim excel, wb, ws
Dim filePath, valueA3, valueA4
Dim resultMsg

' ▼ 引数チェック（ドラッグ&ドロップされたファイルのパス）
If WScript.Arguments.Count = 0 Then
    MsgBox "Excelファイルをドラッグ&ドロップしてください。", vbCritical, "エラー"
    WScript.Quit
End If

filePath = WScript.Arguments(0)

' ▼ Excel起動
Set excel = CreateObject("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False

On Error Resume Next

' ▼ Excelファイルを開く
Set wb = excel.Workbooks.Open(filePath, 0, True)

If Err.Number <> 0 Then
    MsgBox "ファイルを開けませんでした。" & vbCrLf & "エラー: " & Err.Description, vbCritical, "エラー"
    Err.Clear
    On Error GoTo 0
    excel.Quit
    Set excel = Nothing
    WScript.Quit
End If

On Error GoTo 0

' ▼ 1シート目を取得
Set ws = wb.Sheets(1)

' ▼ A3とA4の値を取得
On Error Resume Next
valueA3 = ws.Cells(3, 1).Value2
If Err.Number <> 0 Then
    valueA3 = ""
    Err.Clear
End If

valueA4 = ws.Cells(4, 1).Value2
If Err.Number <> 0 Then
    valueA4 = ""
    Err.Clear
End If
On Error GoTo 0

' ▼ 結果をポップアップで表示
resultMsg = "A3の値: " & CStr(valueA3) & vbCrLf & "A4の値: " & CStr(valueA4)
MsgBox resultMsg, vbInformation, "結果"

' ▼ ファイルを閉じる
wb.Close SaveChanges:=False

' ▼ Excel終了
excel.Quit
Set excel = Nothing
