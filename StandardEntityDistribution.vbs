Option Explicit

'────────────────────────────────────────
'  処理：ドラッグ&ドロップでフォルダを受け取り、
'  フォルダ内のExcelファイルを順次処理し、
'  シート「フィールド」のG7以降をチェックして、
'  「標準」のみが含まれているファイルを
'  「20_標準エンティティ」フォルダに移動
'  ※エラーハンドリングと処理のベースは
'    ExcelFormat.vbsを参考にしています
'────────────────────────────────────────

Dim fso, excel, wb, wsField, ws
Dim folderPath, folder, file
Dim fileName, filePath, fileExt
Dim lastRow, row, gValue
Dim hasCustom, hasStandard
Dim targetFolderPath, targetFolder
Dim isStandardOnly
Dim movedCount
Dim fileList, fileDict
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

' ▼ 移動先フォルダ「20_標準エンティティ」のパスを作成（元のフォルダと同じ階層）
targetFolderPath = fso.GetParentFolderName(folderPath) & "\20_標準エンティティ"

' ▼ 移動先フォルダが存在しない場合は作成
On Error Resume Next
If Not fso.FolderExists(targetFolderPath) Then
    fso.CreateFolder targetFolderPath
    If Err.Number <> 0 Then
        MsgBox "移動先フォルダの作成に失敗しました: " & targetFolderPath & vbCrLf & "エラー: " & Err.Description, vbCritical, "エラー"
        Err.Clear
        WScript.Quit
    End If
End If
On Error GoTo 0

Set targetFolder = fso.GetFolder(targetFolderPath)

' ▼ Excel起動
Set excel = CreateObject("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False
' 高速化設定
excel.ScreenUpdating = False
excel.EnableEvents = False
' Calculation変更は環境によってエラーになるのでエラー無視で実行
On Error Resume Next
excel.Calculation = -4135   ' xlCalculationManual
On Error GoTo 0

movedCount = 0

' ▼ フォルダ内のExcelファイルのリストを先に取得（ループ中にコレクションが変更されても影響を受けないように）
Set fileDict = CreateObject("Scripting.Dictionary")
For Each file In folder.Files
    fileName = file.Name
    fileExt = LCase(fso.GetExtensionName(fileName))
    
    ' Excelファイルの拡張子をチェック
    If fileExt = "xlsx" Or fileExt = "xls" Or fileExt = "xlsm" Or fileExt = "xlsb" Then
        fileDict.Add fileDict.Count, file.Path
    End If
Next

' ▼ 取得したファイルリストを順に処理
For i = 0 To fileDict.Count - 1
    filePath = fileDict(i)
    fileName = fso.GetFileName(filePath)
        
        ' 前のループで作成されたオブジェクトを解放（メモリリーク防止）
        On Error Resume Next
        If Not wsField Is Nothing Then
            Set wsField = Nothing
        End If
        If Not ws Is Nothing Then
            Set ws = Nothing
        End If
        On Error GoTo 0
        
        On Error Resume Next
        ' Excelファイルを開く
        Set wb = excel.Workbooks.Open(filePath, 0, False)
        
        If Err.Number = 0 Then
            On Error GoTo 0
            
            ' シート「フィールド」を取得
            On Error Resume Next
            Set wsField = Nothing
            
            ' シート名で検索
            For Each ws In wb.Sheets
                If ws.Name = "フィールド" Then
                    Set wsField = ws
                    Exit For
                End If
            Next
            
            If Err.Number <> 0 Then
                Err.Clear
            End If
            On Error GoTo 0
            
            ' ▼ シート「フィールド」のG7以降をチェック
            isStandardOnly = True  ' 初期値は「標準のみ」とする
            
            If Not wsField Is Nothing Then
                ' 最終行を取得（G列で判定 - G列に値がある最後の行を取得）
                On Error Resume Next
                lastRow = wsField.Cells(wsField.Rows.Count, 7).End(-4162).Row ' xlUp (G列=7列目)
                If Err.Number <> 0 Or lastRow < 7 Then
                    lastRow = 7
                    Err.Clear
                End If
                On Error GoTo 0
                
                ' G7以降の行をチェック
                If lastRow >= 7 Then
                    For row = 7 To lastRow
                        On Error Resume Next
                        gValue = Trim(CStr(wsField.Cells(row, 7).Value))
                        If Err.Number = 0 And Not IsEmpty(wsField.Cells(row, 7).Value) And gValue <> "" Then
                            ' G列に値がある場合、「カスタム」が含まれているかチェック
                            hasCustom = (InStr(1, gValue, "カスタム", vbTextCompare) > 0)
                            
                            If hasCustom Then
                                ' 「カスタム」が含まれている場合は「標準のみ」ではない
                                isStandardOnly = False
                                Exit For
                            End If
                        End If
                        Err.Clear
                        On Error GoTo 0
                    Next
                Else
                    ' G7以降にデータがない場合は「標準のみ」ではない（判定不能）
                    isStandardOnly = False
                End If
            Else
                ' 「フィールド」シートが見つからない場合は「標準のみ」ではない（判定不能）
                isStandardOnly = False
            End If
            
            ' ファイルを閉じる（保存しない）
            On Error Resume Next
            wb.Close False
            If Err.Number <> 0 Then
                Err.Clear
            End If
            On Error GoTo 0
            
            Set wb = Nothing
            Set wsField = Nothing
            Set ws = Nothing
            
            ' ▼ 「標準のみ」のファイルを移動
            If isStandardOnly Then
                On Error Resume Next
                Set file = fso.GetFile(filePath)
                file.Move targetFolderPath & "\" & fileName
                If Err.Number = 0 Then
                    movedCount = movedCount + 1
                Else
                    MsgBox "ファイルの移動に失敗しました: " & fileName & vbCrLf & "エラー: " & Err.Description, vbCritical, "エラー"
                    Err.Clear
                End If
                On Error GoTo 0
                Set file = Nothing
            End If
        Else
            MsgBox "ファイルを開けませんでした: " & fileName & vbCrLf & "エラー: " & Err.Description, vbCritical, "エラー"
            Err.Clear
            On Error GoTo 0
        End If
Next

' ▼ Excel終了（設定を戻してから Quit）
On Error Resume Next
excel.Calculation = -4105   ' xlCalculationAutomatic（失敗しても無視）
excel.ScreenUpdating = True
excel.EnableEvents = True
On Error GoTo 0

excel.Quit
Set excel = Nothing

' オブジェクトを解放
Set fileDict = Nothing
Set fso = Nothing
Set folder = Nothing
Set targetFolder = Nothing

MsgBox "処理が完了しました。" & vbCrLf & "移動したファイル数: " & movedCount, vbInformation, "完了"

