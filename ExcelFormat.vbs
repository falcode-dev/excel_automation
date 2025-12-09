Option Explicit

'────────────────────────────────────────
'  処理：ドラッグ&ドロップでフォルダを受け取り、
'  フォルダ内のExcelファイルを順次処理し、
'  シート「テーブル」のE5からE41を黒文字に、
'  シート「フィールド」のD7からAK417を黒文字にし、
'  それぞれA1にカーソルを戻してシートの1枚目にして保存
'  ※高速化ポイント
'    ・ScreenUpdating、EnableEvents、Calculationを無効化
'    ・エラーハンドリングを最適化
'    ・オブジェクト解放を確実に実行
'────────────────────────────────────────

Dim fso, excel, wb, wsTable, wsField, ws, wsCover
Dim folderPath, folder, file
Dim fileName, filePath, fileExt
Dim lastRow, row, gValue, cValue, dValue, lowerGValue
Dim hasCustom, hasStandard
Dim dataArr, sortedArr, customRows, standardRows, emptyRows
Dim i, j, arrRow, colCount, startRow, startCol
Dim rowData
Dim emptyStartRow, deleteEndRow, checkRow
Dim dRowCount, deleteStartRow
Dim dColArr, dMaxRow

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
        If Not wsCover Is Nothing Then
            Set wsCover = Nothing
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
            
            ' シート「テーブル」「フィールド」「表紙」を取得
            On Error Resume Next
            Set wsTable = Nothing
            Set wsField = Nothing
            Set wsCover = Nothing
            
            ' シート名で検索
            For Each ws In wb.Sheets
                If ws.Name = "テーブル" Then
                    Set wsTable = ws
                ElseIf ws.Name = "フィールド" Then
                    Set wsField = ws
                ElseIf ws.Name = "表紙" Then
                    Set wsCover = ws
                End If
            Next
            
            If Err.Number <> 0 Then
                Err.Clear
            End If
            On Error GoTo 0
            
            ' ▼ シート「テーブル」の処理
            If Not wsTable Is Nothing Then
                ' E5からE41を黒文字にする
                With wsTable.Range("E5:E41")
                    .Font.Color = RGB(0, 0, 0)  ' 黒文字
                End With
            End If
            
            ' ▼ シート「フィールド」の処理
            If Not wsField Is Nothing Then
                ' D7からAK417を黒文字にする
                With wsField.Range("D7:AK417")
                    .Font.Color = RGB(0, 0, 0)  ' 黒文字
                End With
                
                ' ▼ G7以降の行を処理（並び替え、削除、C列への値設定）
                ' 最終行を取得（AK列で判定）
                On Error Resume Next
                lastRow = wsField.Cells(wsField.Rows.Count, 37).End(-4162).Row ' xlUp (AK列=37列目)
                If Err.Number <> 0 Or lastRow < 7 Then
                    lastRow = 7
                    Err.Clear
                End If
                On Error GoTo 0
                
                If lastRow >= 7 Then
                    startRow = 7
                    startCol = 1  ' A列から
                    colCount = 37 ' AK列まで（37列目）
                    
                    ' 7行目以降のデータを一括で配列に読み込む（高速化）
                    On Error Resume Next
                    dataArr = wsField.Range(wsField.Cells(startRow, startCol), wsField.Cells(lastRow, colCount)).Value
                    If Err.Number <> 0 Then
                        dataArr = Array()
                        Err.Clear
                    End If
                    On Error GoTo 0
                    
                    ' データを分類（カスタム、標準、空）
                    Set customRows = CreateObject("Scripting.Dictionary")
                    Set standardRows = CreateObject("Scripting.Dictionary")
                    Set emptyRows = CreateObject("Scripting.Dictionary")
                    
                    ' 配列は1ベース（Excelから読み込んだ場合）
                    If IsArray(dataArr) Then
                        Dim maxRow, maxCol
                        On Error Resume Next
                        maxRow = UBound(dataArr, 1)
                        maxCol = UBound(dataArr, 2)
                        If Err.Number <> 0 Then
                            maxRow = 0
                            maxCol = 0
                            Err.Clear
                        End If
                        On Error GoTo 0
                        
                        arrRow = 0
                        For row = startRow To lastRow
                            On Error Resume Next
                            gValue = ""
                            If maxCol >= 7 And arrRow + 1 <= maxRow Then  ' G列は7列目
                                gValue = Trim(CStr(dataArr(arrRow + 1, 7)))  ' G列（7列目）
                                If Err.Number <> 0 Or IsEmpty(dataArr(arrRow + 1, 7)) Then
                                    gValue = ""
                                    Err.Clear
                                End If
                            End If
                            On Error GoTo 0
                            
                            ' G列の値に基づいて分類
                            If gValue = "" Or IsEmpty(gValue) Then
                                ' G列が空の行は削除対象
                                emptyRows.Add arrRow, row
                            Else
                                ' G列に値がある場合、C列に "proto_" をセット（既存の値がない場合のみ）
                                On Error Resume Next
                                If maxCol >= 3 And arrRow + 1 <= maxRow Then  ' C列は3列目
                                    cValue = Trim(CStr(dataArr(arrRow + 1, 3)))
                                    If Err.Number <> 0 Or cValue = "" Or IsEmpty(dataArr(arrRow + 1, 3)) Then
                                        dataArr(arrRow + 1, 3) = "proto_"
                                    End If
                                    Err.Clear
                                End If
                                On Error GoTo 0
                                
                                ' 「カスタム」と「標準」で分類（部分一致にも対応）
                                ' G列の値に「カスタム」が含まれている場合はカスタムとして扱う
                                ' 「標準」が含まれている場合は標準として扱う
                                ' 両方含まれている場合はカスタムを優先
                                ' InStrは大文字小文字を区別しないが、日本語には影響しないため、そのまま使用
                                hasCustom = (InStr(1, gValue, "カスタム", vbTextCompare) > 0)
                                hasStandard = (InStr(1, gValue, "標準", vbTextCompare) > 0)
                                
                                If hasCustom Then
                                    ' 「カスタム」が含まれている場合はカスタムとして扱う
                                    customRows.Add customRows.Count, arrRow
                                ElseIf hasStandard Then
                                    ' 「標準」が含まれている場合は標準として扱う
                                    standardRows.Add standardRows.Count, arrRow
                                Else
                                    ' その他の値も標準として扱う
                                    standardRows.Add standardRows.Count, arrRow
                                End If
                            End If
                            arrRow = arrRow + 1
                        Next
                    End If
                    
                    ' 並び替えたデータを作成（カスタム→標準の順）
                    Dim sortedCount, totalRows
                    totalRows = customRows.Count + standardRows.Count
                    
                    If totalRows > 0 Then
                        ' Excelに書き込む配列は0ベースで作成（Excelはインデックス0を1行目として解釈）
                        ReDim sortedArr(totalRows - 1, colCount - 1)
                        sortedCount = 0  ' 0ベースで開始
                        
                        ' カスタムの行を追加
                        For i = 0 To customRows.Count - 1
                            arrRow = customRows(i)
                            For j = 0 To colCount - 1
                                On Error Resume Next
                                ' dataArrは1ベース（Excelから読み込んだ）、sortedArrは0ベース
                                sortedArr(sortedCount, j) = dataArr(arrRow + 1, j + 1)
                                If Err.Number <> 0 Then
                                    sortedArr(sortedCount, j) = ""
                                    Err.Clear
                                End If
                                On Error GoTo 0
                            Next
                            sortedCount = sortedCount + 1
                        Next
                        
                        ' 標準の行を追加
                        For i = 0 To standardRows.Count - 1
                            arrRow = standardRows(i)
                            For j = 0 To colCount - 1
                                On Error Resume Next
                                ' dataArrは1ベース（Excelから読み込んだ）、sortedArrは0ベース
                                sortedArr(sortedCount, j) = dataArr(arrRow + 1, j + 1)
                                If Err.Number <> 0 Then
                                    sortedArr(sortedCount, j) = ""
                                    Err.Clear
                                End If
                                On Error GoTo 0
                            Next
                            sortedCount = sortedCount + 1
                        Next
                        
                        ' 既存のデータをクリア（7行目以降）
                        If lastRow >= startRow Then
                            wsField.Range(wsField.Cells(startRow, startCol), wsField.Cells(lastRow, colCount)).ClearContents
                        End If
                        
                        ' 並び替えたデータを書き戻す
                        ' Excelは配列のインデックス0を1行目として解釈する
                        wsField.Range(wsField.Cells(startRow, startCol), wsField.Cells(startRow + totalRows - 1, colCount)).Value = sortedArr
                    Else
                        ' データがない場合は7行目以降をクリア
                        If lastRow >= startRow Then
                            wsField.Range(wsField.Cells(startRow, startCol), wsField.Cells(lastRow, colCount)).ClearContents
                        End If
                    End If
                    
                    ' オブジェクトを解放
                    Set customRows = Nothing
                    Set standardRows = Nothing
                    Set emptyRows = Nothing
                    dataArr = Array()
                    sortedArr = Array()
                    
                    ' ▼ D7以降の値のある行数を取得し、その次の行から削除
                    ' 並び替え後の最終行を再取得（D列で判定）
                    On Error Resume Next
                    lastRow = wsField.Cells(wsField.Rows.Count, 4).End(-4162).Row ' xlUp (D列=4列目)
                    If Err.Number <> 0 Or lastRow < 7 Then
                        lastRow = 7
                        Err.Clear
                    End If
                    On Error GoTo 0
                    
                    ' D7以降で値がある行数をカウント（配列で一括読み込みして高速化）
                    dRowCount = 0
                    If lastRow >= 7 Then
                        ' D列の7行目以降を配列で一括読み込み
                        On Error Resume Next
                        dColArr = wsField.Range(wsField.Cells(7, 4), wsField.Cells(lastRow, 4)).Value
                        If Err.Number = 0 And IsArray(dColArr) Then
                            dMaxRow = UBound(dColArr, 1)
                            For checkRow = 1 To dMaxRow
                                On Error Resume Next
                                dValue = Trim(CStr(dColArr(checkRow, 1)))
                                If Err.Number = 0 And Not IsEmpty(dColArr(checkRow, 1)) And dValue <> "" Then
                                    ' 値がある行をカウント
                                    dRowCount = dRowCount + 1
                                End If
                                Err.Clear
                                On Error GoTo 0
                            Next
                        End If
                        Err.Clear
                        On Error GoTo 0
                        dColArr = Array() ' メモリ解放
                    End If
                    
                    ' D7を含めた行（7行目）+ 値のある行数 + 1行目から削除
                    ' 例：D7からD10まで値があれば、7 + 4 + 1 = 12行目から削除
                    If dRowCount > 0 Then
                        deleteStartRow = 7 + dRowCount
                        
                        ' 削除する最終行を計算（シートの最終行まで）
                        deleteEndRow = wsField.Rows.Count
                        
                        ' 削除開始行がシートの最終行を超えない場合のみ削除
                        If deleteStartRow <= deleteEndRow Then
                            ' 行を削除
                            On Error Resume Next
                            wsField.Rows(deleteStartRow & ":" & deleteEndRow).Delete
                            If Err.Number <> 0 Then
                                Err.Clear
                            End If
                            On Error GoTo 0
                        End If
                    End If
                    
                    ' ▼ B7以降に =ROW()-6 をセット（一括設定で高速化）
                    ' 削除後の最終行を再取得（D列で判定）
                    On Error Resume Next
                    lastRow = wsField.Cells(wsField.Rows.Count, 4).End(-4162).Row ' xlUp (D列=4列目)
                    If Err.Number <> 0 Or lastRow < 7 Then
                        lastRow = 7
                        Err.Clear
                    End If
                    On Error GoTo 0
                    
                    ' B7から最終行まで =ROW()-6 を一括でセット（ループではなくRangeで一括設定）
                    If lastRow >= 7 Then
                        On Error Resume Next
                        wsField.Range(wsField.Cells(7, 2), wsField.Cells(lastRow, 2)).Formula = "=ROW()-6"
                        If Err.Number <> 0 Then
                            Err.Clear
                        End If
                        On Error GoTo 0
                    End If
                    
                    ' ▼ J列からW列の幅を自動調整（J列=10列目、W列=23列目）
                    On Error Resume Next
                    wsField.Range(wsField.Cells(1, 10), wsField.Cells(1, 23)).Columns.AutoFit
                    If Err.Number <> 0 Then
                        Err.Clear
                    End If
                    On Error GoTo 0
                    
                    ' ▼ 列の幅を個別に設定
                    On Error Resume Next
                    wsField.Columns(10).ColumnWidth = 27  ' J列：27
                    wsField.Columns(11).ColumnWidth = 15  ' K列：15
                    wsField.Columns(12).ColumnWidth = 15  ' L列：15
                    wsField.Columns(21).ColumnWidth = 15  ' U列：15
                    ' T列（20列目）を折り返しに設定
                    wsField.Columns(20).WrapText = True
                    If Err.Number <> 0 Then
                        Err.Clear
                    End If
                    On Error GoTo 0
                End If
            End If
            
            ' シートが見つからない場合の警告
            If wsTable Is Nothing And wsField Is Nothing Then
                MsgBox "シート「テーブル」または「フィールド」が見つかりません: " & fileName, vbWarning, "警告"
            ElseIf wsTable Is Nothing Then
                MsgBox "シート「テーブル」が見つかりません: " & fileName, vbWarning, "警告"
            ElseIf wsField Is Nothing Then
                MsgBox "シート「フィールド」が見つかりません: " & fileName, vbWarning, "警告"
            End If
            
            ' 最後に「表紙」シートをアクティブにしてA1にカーソルをセット（1枚目にする）
            If Not wsCover Is Nothing Then
                wsCover.Activate
                wsCover.Range("A1").Select
            ElseIf Not wsTable Is Nothing Then
                wsTable.Activate
                wsTable.Range("A1").Select
            ElseIf Not wsField Is Nothing Then
                wsField.Activate
                wsField.Range("A1").Select
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
            ' 保存ダイアログを防ぐため、SavedプロパティをTrueに設定
            wb.Saved = True
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
            Set wsCover = Nothing
            Set ws = Nothing
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

MsgBox "処理が完了しました。", vbInformation, "完了"

