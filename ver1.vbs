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
Dim i, filePath, cell3, cell4, val3, val4
Dim errorMsg, nonExcelFile, openErrMsg
Dim processedCount, errOccurred
Dim wbTemp

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
    errorMsg = "フォルダ内にExcel以外のファイルが含まれています。" & vbCrLf & vbCrLf
    errorMsg = errorMsg & "Excel以外のファイル:" & vbCrLf
    
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

' ▼ Excel起動（日本語対応設定）
Set excel = CreateObject("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False
' 日本語環境に対応
excel.EnableEvents = False
excel.ScreenUpdating = False

' ▼ 各Excelファイルを処理
processedCount = 0
errOccurred = False

For Each fileName In excelFiles.Keys
    ' 各ファイル処理前にwbを初期化
    Set wb = Nothing
    
    On Error Resume Next
    
    ' Excelファイルを開く（日本語パス対応）
    filePath = excelFiles(fileName)
    ' パスを正しく処理（日本語を含む場合も対応）
    ' VBSでは名前付き引数は使えないため、位置引数で指定
    ' Workbooks.Open(FileName, UpdateLinks, ReadOnly, ...)
    Set wb = excel.Workbooks.Open(filePath, 0, True)
    
    If Err.Number <> 0 Then
        openErrMsg = "ファイルを開けませんでした: " & fileName & vbCrLf & "エラー: " & Err.Description
        MsgBox openErrMsg, vbCritical, "エラー"
        Err.Clear
        On Error GoTo 0
        ' ファイルが開けなかった場合は次のファイルへ
        GoTo CLEANUP_FILE
    End If
    
    On Error GoTo 0
    
    ' 1シート目を取得
    On Error Resume Next
    If wb.Sheets.Count = 0 Then
        MsgBox "シートが存在しません: " & fileName, vbWarning, "警告"
        On Error GoTo 0
        GoTo CLEANUP_FILE
    End If
    
    Set ws = wb.Sheets(1)
    
    If Err.Number <> 0 Then
        MsgBox "シートの取得に失敗しました: " & fileName & vbCrLf & "エラー: " & Err.Description, vbCritical, "エラー"
        Err.Clear
        On Error GoTo 0
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
        
        ' 3行目の値を取得（日本語対応：Textプロパティを使用）
        Set cell3 = ws.Cells(3, col)
        Set cell4 = ws.Cells(4, col)
        
        If Not IsEmpty(cell3.Value) Then
            ' 日本語対応：Value2プロパティを優先的に使用（文字化けを防ぐ）
            On Error Resume Next
            val3 = cell3.Value2
            If Err.Number = 0 Then
                ' Value2が取得できた場合
                If IsNumeric(val3) Then
                    cellValue3 = CStr(val3)
                Else
                    ' 文字列の場合はValue2をそのまま使用（日本語対応）
                    cellValue3 = CStr(val3)
                End If
            Else
                ' Value2が取得できない場合はValueを使用
                Err.Clear
                cellValue3 = CStr(cell3.Value)
            End If
            On Error GoTo 0
        End If
        
        ' 4行目の値を取得（日本語対応：Value2プロパティを優先的に使用）
        If Not IsEmpty(cell4.Value) Then
            On Error Resume Next
            val4 = cell4.Value2
            If Err.Number = 0 Then
                ' Value2が取得できた場合
                If IsNumeric(val4) Then
                    cellValue4 = CStr(val4)
                Else
                    ' 文字列の場合はValue2をそのまま使用（日本語対応）
                    cellValue4 = CStr(val4)
                End If
            Else
                ' Value2が取得できない場合はValueを使用
                Err.Clear
                cellValue4 = CStr(cell4.Value)
            End If
            On Error GoTo 0
        End If
        
        ' 紐づけて表示（列番号: 3行目の値 → 4行目の値）
        resultMsg = resultMsg & "列" & col & ": " & cellValue3 & " → " & cellValue4 & vbCrLf
        
        ' エラーが発生した場合はループを抜ける
        If Err.Number <> 0 Then
            Err.Clear
            Exit For
        End If
        
        Set cell3 = Nothing
        Set cell4 = Nothing
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
    For Each wbTemp In excel.Workbooks
        wbTemp.Close SaveChanges:=False
    Next
    
    excel.ScreenUpdating = True
    excel.EnableEvents = True
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

