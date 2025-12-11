Option Explicit

'────────────────────────────────────────
'  処理：ドラッグ&ドロップでフォルダを受け取り、
'  フォルダ内のExcelファイルを順次処理し、
'  ・シート「テーブル」のE5からE21の範囲で
'    tel_wo_ を wo_ に置換
'  ・シート「フィールド」のD7以降で
'    tel_wo_ を wo_ に置換
'  ※エラーハンドリングと処理のベースは
'    ItemConversion.vbsを参考にしています
'────────────────────────────────────────

Dim fso, excel, wb, wsTable, wsField, ws
Dim folderPath, folder, file
Dim fileName, filePath, fileExt
Dim linkSources, linkIndex, linkCount
Dim lastRow
Dim cellRange, cellValue, i, j
Dim dataArray, rowCount, colCount

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
' 高速化設定
excel.ScreenUpdating = False
excel.EnableEvents = False
' Calculation変更は環境によってエラーになるのでエラー無視で実行
On Error Resume Next
excel.Calculation = -4135   ' xlCalculationManual
On Error GoTo 0

' ▼ フォルダ内のExcelファイルを順に処理
For Each file In folder.Files
    fileName = file.Name
    fileExt = LCase(fso.GetExtensionName(fileName))
    
    ' Excelファイルの拡張子をチェック
    If fileExt = "xlsx" Or fileExt = "xls" Or fileExt = "xlsm" Or fileExt = "xlsb" Then
        filePath = file.Path
        
        ' 前のループで作成されたオブジェクトを解放（メモリリーク防止）
        On Error Resume Next
        If Not wsTable Is Nothing Then
            Set wsTable = Nothing
        End If
        If Not wsField Is Nothing Then
            Set wsField = Nothing
        End If
        If Not ws Is Nothing Then
            Set ws = Nothing
        End If
        On Error GoTo 0
        
        On Error Resume Next
        ' Excelファイルを開く（リンクの更新を無効にする）
        Set wb = excel.Workbooks.Open(filePath, 0, False)
        
        If Err.Number = 0 Then
            On Error GoTo 0
            ' リンクの更新を無効にする（警告を表示しない）
            On Error Resume Next
            wb.UpdateLinks = 0  ' xlUpdateLinksNever
            If Err.Number <> 0 Then
                Err.Clear
            End If
            On Error GoTo 0
            
            ' シート「テーブル」を取得
            On Error Resume Next
            Set wsTable = Nothing
            
            ' シート名で検索
            For Each ws In wb.Sheets
                If ws.Name = "テーブル" Then
                    Set wsTable = ws
                    Exit For
                End If
            Next
            
            If Err.Number <> 0 Then
                Err.Clear
            End If
            On Error GoTo 0
            
            ' ▼ シート「テーブル」のE5からE21の範囲で tel_wo_ を wo_ に置換
            If Not wsTable Is Nothing Then
                On Error Resume Next
                ' E5からE21の範囲を取得（5列目、5行目から21行目）
                Set cellRange = wsTable.Range(wsTable.Cells(5, 5), wsTable.Cells(21, 5))
                
                ' 範囲を配列として取得（高速化）
                dataArray = cellRange.Value
                
                ' 配列が2次元配列の場合（複数セル）
                If IsArray(dataArray) Then
                    rowCount = UBound(dataArray, 1)
                    colCount = UBound(dataArray, 2)
                    
                    ' 各セルの値をチェックして置換
                    For i = 1 To rowCount
                        For j = 1 To colCount
                            If Not IsEmpty(dataArray(i, j)) Then
                                cellValue = CStr(dataArray(i, j))
                                ' tel_wo_ を wo_ に置換
                                If InStr(cellValue, "tel_wo_") > 0 Then
                                    dataArray(i, j) = Replace(cellValue, "tel_wo_", "wo_")
                                End If
                            End If
                        Next
                    Next
                    
                    ' 置換後の値を一括で書き戻し（高速化）
                    cellRange.Value = dataArray
                Else
                    ' 単一セルの場合
                    If Not IsEmpty(dataArray) Then
                        cellValue = CStr(dataArray)
                        If InStr(cellValue, "tel_wo_") > 0 Then
                            cellRange.Value = Replace(cellValue, "tel_wo_", "wo_")
                        End If
                    End If
                End If
                
                Err.Clear
                On Error GoTo 0
            End If
            
            ' ▼ シート「フィールド」の処理
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
            
            ' ▼ シート「フィールド」のD7以降で tel_wo_ を wo_ に置換
            If Not wsField Is Nothing Then
                On Error Resume Next
                ' D列の最後の行を取得（高速化のため、End(xlUp)を使用）
                lastRow = wsField.Cells(wsField.Rows.Count, 4).End(-4162).Row  ' xlUp = -4162
                
                ' 7行目以降にデータがある場合のみ処理
                If lastRow >= 7 Then
                    ' D7からD列の最後の行までの範囲を取得
                    Set cellRange = wsField.Range(wsField.Cells(7, 4), wsField.Cells(lastRow, 4))
                    
                    ' 範囲を配列として取得（高速化）
                    dataArray = cellRange.Value
                    
                    ' 配列が2次元配列の場合（複数セル）
                    If IsArray(dataArray) Then
                        rowCount = UBound(dataArray, 1)
                        colCount = UBound(dataArray, 2)
                        
                        ' 各セルの値をチェックして置換
                        For i = 1 To rowCount
                            For j = 1 To colCount
                                If Not IsEmpty(dataArray(i, j)) Then
                                    cellValue = CStr(dataArray(i, j))
                                    ' tel_wo_ を wo_ に置換
                                    If InStr(cellValue, "tel_wo_") > 0 Then
                                        dataArray(i, j) = Replace(cellValue, "tel_wo_", "wo_")
                                    End If
                                End If
                            Next
                        Next
                        
                        ' 置換後の値を一括で書き戻し（高速化）
                        cellRange.Value = dataArray
                    Else
                        ' 単一セルの場合
                        If Not IsEmpty(dataArray) Then
                            cellValue = CStr(dataArray)
                            If InStr(cellValue, "tel_wo_") > 0 Then
                                cellRange.Value = Replace(cellValue, "tel_wo_", "wo_")
                            End If
                        End If
                    End If
                End If
                
                Err.Clear
                On Error GoTo 0
            End If
            
            ' ▼ 「保存しますか？」ダイアログを防ぐため、保存前に計算を実行
            ' 計算モードを自動に戻してから計算を実行
            ' 注意：循環参照がある場合は計算が終わらない可能性がある
            ' 高速化のため、計算モード変更と計算を最後に一度だけ実行
            On Error Resume Next
            excel.Calculation = -4105   ' xlCalculationAutomatic
            ' 計算を実行（循環参照がある場合は時間がかかる可能性がある）
            ' Calculateは同期的に実行されるため、計算が完了するまで待機する
            wb.Calculate   ' ブック全体を計算
            If Err.Number <> 0 Then
                ' 計算エラーが発生した場合（循環参照など）は計算モードを手動に戻す
                excel.Calculation = -4135   ' xlCalculationManual
                Err.Clear
            End If
            
            ' ▼ 外部リンクを削除して警告を防ぐ
            On Error Resume Next
            linkSources = wb.LinkSources(1)  ' xlExcelLinks
            If Err.Number = 0 Then
                ' リンクが存在する場合
                If IsArray(linkSources) Then
                    On Error Resume Next
                    linkCount = UBound(linkSources) + 1
                    If Err.Number = 0 And linkCount > 0 Then
                        For linkIndex = 0 To linkCount - 1
                            On Error Resume Next
                            wb.BreakLink linkSources(linkIndex), 1  ' xlLinkTypeExcelLinks
                            Err.Clear
                            On Error GoTo 0
                        Next
                    End If
                    Err.Clear
                ElseIf Not IsEmpty(linkSources) And linkSources <> "" Then
                    ' リンクが1つの場合（配列ではなく文字列）
                    On Error Resume Next
                    wb.BreakLink linkSources, 1  ' xlLinkTypeExcelLinks
                    Err.Clear
                End If
            End If
            Err.Clear
            On Error GoTo 0
            
            ' 保存ダイアログを防ぐため、SavedプロパティをTrueに設定
            wb.Saved = True
            ' リンクの更新を無効にする（保存時の警告を防ぐ）
            On Error Resume Next
            wb.UpdateLinks = 0  ' xlUpdateLinksNever
            If Err.Number <> 0 Then
                Err.Clear
            End If
            On Error GoTo 0
            
            ' ファイルを保存
            On Error Resume Next
            wb.Save
            If Err.Number <> 0 Then
                MsgBox "ファイルの保存に失敗しました: " & fileName & vbCrLf & "エラー: " & Err.Description, vbCritical, "エラー"
                Err.Clear
            End If
            On Error GoTo 0
            
            ' ファイルを閉じる
            wb.Close False
            Set wb = Nothing
            Set wsTable = Nothing
            Set wsField = Nothing
            Set ws = Nothing
            Set cellRange = Nothing
        Else
            MsgBox "ファイルを開けませんでした: " & fileName & vbCrLf & "エラー: " & Err.Description, vbCritical, "エラー"
            Err.Clear
            On Error GoTo 0
        End If
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
Set fso = Nothing
Set folder = Nothing

MsgBox "処理が完了しました。", vbInformation, "完了"
