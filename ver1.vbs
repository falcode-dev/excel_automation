Option Explicit

'────────────────────────────────────────
'  シンプル処理：ドラッグ&ドロップでフォルダを受け取り、
'  フォルダ内のExcelファイルのA3とA4の値を順にポップアップで表示
'────────────────────────────────────────

Dim fso, excel, wb, ws
Dim folderPath, folder, file
Dim fileName, filePath, fileExt
Dim valueA3, valueA4, resultMsg

' ▼ 引数チェック（ドラッグ&ドロップされたフォルダのパス）
If WScript.Arguments.Count = 0 Then
    MsgBox "フォルダをドラッグ&ドロップしてください。", vbCritical, "エラー"
    WScript.Quit
End If

folderPath = WScript.Arguments(0)

' ▼ FileSystemObject作成
Set fso = CreateObject("Scripting.FileSystemObject")

' ▼ フォルダの存在チェック
If Not fso.FolderExists(folderPath) Then
    MsgBox "指定されたパスはフォルダではありません: " & folderPath, vbCritical, "エラー"
    WScript.Quit
End If

Set folder = fso.GetFolder(folderPath)

' ▼ Excel起動
Set excel = CreateObject("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False

' ▼ フォルダ内のExcelファイルを順に処理
For Each file In folder.Files
    fileName = file.Name
    fileExt = LCase(fso.GetExtensionName(fileName))
    
    ' Excelファイルの拡張子をチェック
    If fileExt = "xlsx" Or fileExt = "xls" Or fileExt = "xlsm" Or fileExt = "xlsb" Then
        filePath = file.Path
        
        On Error Resume Next
        
        ' Excelファイルを開く
        Set wb = excel.Workbooks.Open(filePath, 0, True)
        
        If Err.Number <> 0 Then
            MsgBox "ファイルを開けませんでした: " & fileName & vbCrLf & "エラー: " & Err.Description, vbCritical, "エラー"
            Err.Clear
            On Error GoTo 0
            GoTo NEXT_FILE
        End If
        
        On Error GoTo 0
        
        ' 1シート目を取得
        Set ws = wb.Sheets(1)
        
        ' A3とA4の値を取得
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
        
        ' 結果をポップアップで表示
        resultMsg = "ファイル: " & fileName & vbCrLf & vbCrLf
        resultMsg = resultMsg & "A3の値: " & CStr(valueA3) & vbCrLf & "A4の値: " & CStr(valueA4)
        MsgBox resultMsg, vbInformation, "結果"
        
        ' ファイルを閉じる
        wb.Close False
        Set wb = Nothing
        Set ws = Nothing
        
NEXT_FILE:
    End If
Next

' ▼ Excel終了
excel.Quit
Set excel = Nothing
