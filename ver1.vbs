Option Explicit

'────────────────────────────────────────
'  メイン処理：ドラッグ&ドロップでフォルダを受け取り、
'  フォルダ内のExcelファイルの3行目と4行目を紐づけて表示
'────────────────────────────────────────

Dim fso, excel, folderPath, folder
Dim file, fileName, fileExt
Dim excelFiles, nonExcelFiles
Dim wb, ws
Dim row3Data, row4Data
Dim lastCol, col
Dim resultMsg, cellValue3, cellValue4
Dim i

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

' ▼ フォルダ内のファイルをチェック
Set excelFiles = CreateObject("Scripting.Dictionary")
Set nonExcelFiles = CreateObject("Scripting.Dictionary")

For Each file In folder.Files
    fileName = file.Name
    fileExt = LCase(fso.GetExtensionName(fileName))
    
    ' Excelファイルの拡張子をチェック
    If fileExt = "xlsx" Or fileExt = "xls" Or fileExt = "xlsm" Or fileExt = "xlsb" Then
        excelFiles.Add fileName, file.Path
    Else
        nonExcelFiles.Add fileName, fileExt
    End If
Next

' ▼ Excel以外のファイルがある場合はエラー
If nonExcelFiles.Count > 0 Then
    Dim errorMsg
    errorMsg = "フォルダ内にExcel以外のファイルが含まれています。" & vbCrLf & vbCrLf
    errorMsg = errorMsg & "Excel以外のファイル:" & vbCrLf
    
    Dim nonExcelFile
    For Each nonExcelFile In nonExcelFiles.Keys
        errorMsg = errorMsg & "  - " & nonExcelFile & " (" & nonExcelFiles(nonExcelFile) & ")" & vbCrLf
    Next
    
    errorMsg = errorMsg & vbCrLf & "処理をキャンセルします。"
    MsgBox errorMsg, vbCritical, "エラー"
    WScript.Quit
End If

' ▼ Excelファイルがない場合
If excelFiles.Count = 0 Then
    MsgBox "フォルダ内にExcelファイルが見つかりませんでした。", vbCritical, "エラー"
    WScript.Quit
End If

' ▼ Excel起動
Set excel = CreateObject("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False

' ▼ 各Excelファイルを処理
Dim processedCount
processedCount = 0

' ▼ エラーハンドリング用の変数
Dim errOccurred
errOccurred = False

For Each fileName In excelFiles.Keys
    ' 各ファイル処理前にwbを初期化
    Set wb = Nothing
    
    On Error Resume Next
    
    ' Excelファイルを開く
    Set wb = excel.Workbooks.Open(excelFiles(fileName), ReadOnly:=True)
    
    If Err.Number <> 0 Then
        Dim openErrMsg
        openErrMsg = "ファイルを開けませんでした: " & fileName & vbCrLf & "エラー: " & Err.Description
        MsgBox openErrMsg, vbCritical, "エラー"
        Err.Clear
        ' ファイルが開けなかった場合は次のファイルへ
        GoTo CLEANUP_FILE
    End If
    
    On Error GoTo 0
    
    ' 1シート目を取得
    On Error Resume Next
    If wb.Sheets.Count = 0 Then
        MsgBox "シートが存在しません: " & fileName, vbWarning, "警告"
        GoTo CLEANUP_FILE
    End If
    
    Set ws = wb.Sheets(1)
    
    If Err.Number <> 0 Then
        MsgBox "シートの取得に失敗しました: " & fileName & vbCrLf & "エラー: " & Err.Description, vbCritical, "エラー"
        Err.Clear
        GoTo CLEANUP_FILE
    End If
    
    On Error GoTo 0
    
    ' 最終列を取得
    On Error Resume Next
    lastCol = ws.Cells(3, ws.Columns.Count).End(-4159).Column ' xlToLeft
    
    If Err.Number <> 0 Then
        ' 最終列の取得に失敗した場合は1列目から処理
        lastCol = 1
        Err.Clear
    End If
    
    On Error GoTo 0
    
    ' 3行目と4行目のデータを取得して紐づける
    resultMsg = "ファイル: " & fileName & vbCrLf & vbCrLf
    resultMsg = resultMsg & "【3行目と4行目の紐づけ】" & vbCrLf & vbCrLf
    
    On Error Resume Next
    For col = 1 To lastCol
        cellValue3 = ""
        cellValue4 = ""
        
        ' 3行目の値を取得
        If Not IsEmpty(ws.Cells(3, col).Value) Then
            cellValue3 = CStr(ws.Cells(3, col).Value)
        End If
        
        ' 4行目の値を取得
        If Not IsEmpty(ws.Cells(4, col).Value) Then
            cellValue4 = CStr(ws.Cells(4, col).Value)
        End If
        
        ' 紐づけて表示（列番号: 3行目の値 → 4行目の値）
        resultMsg = resultMsg & "列" & col & ": " & cellValue3 & " → " & cellValue4 & vbCrLf
        
        ' エラーが発生した場合はループを抜ける
        If Err.Number <> 0 Then
            Err.Clear
            Exit For
        End If
    Next col
    On Error GoTo 0
    
    ' 結果を表示
    MsgBox resultMsg, vbInformation, "処理結果: " & fileName
    
    processedCount = processedCount + 1
    
CLEANUP_FILE:
    ' ▼ ファイルを確実に閉じる（エラーが発生しても実行）
    On Error Resume Next
    
    If Not wb Is Nothing Then
        wb.Close SaveChanges:=False
    End If
    
    Set wb = Nothing
    Set ws = Nothing
    
    ' エラーをクリア
    Err.Clear
    On Error GoTo 0
    
Next

' ▼ Excel終了（エラーが発生しても確実に終了させる）
On Error Resume Next
If Not excel Is Nothing Then
    ' 開いているワークブックがあればすべて閉じる
    Dim wbTemp
    For Each wbTemp In excel.Workbooks
        wbTemp.Close SaveChanges:=False
    Next
    
    excel.Quit
End If
Set excel = Nothing
Err.Clear
On Error GoTo 0

' ▼ 完了メッセージ
If processedCount > 0 Then
    MsgBox "処理が完了しました。" & vbCrLf & "処理したファイル数: " & processedCount, vbInformation, "完了"
Else
    MsgBox "処理できたファイルがありませんでした。", vbWarning, "警告"
End If

