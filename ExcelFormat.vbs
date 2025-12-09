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

Dim fso, excel, wb, wsTable, wsField, wsCover, ws
Dim folderPath, folder, file
Dim fileName, filePath, fileExt
Dim lastRow, row, gValue, cValue
Dim dataArr, sortedArr, customRows, standardRows, emptyRows
Dim i, j, arrRow, colCount, startRow, startCol
Dim rowData

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
        ' Excelファイルを開く
        Set wb = excel.Workbooks.Open(filePath, 0, False)
        
        If Err.Number = 0 Then
            On Error GoTo 0
            
            ' エラー発生時のクリーンアップ用にフラグを設定
            Dim fileProcessed
            fileProcessed = False
            
            On Error Resume Next
            ' シート「テーブル」「フィールド」「表紙」を取得
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
            On Error Resume Next
            If Not wsTable Is Nothing Then
                ' E5からE41を黒文字にする
                With wsTable.Range("E5:E41")
                    .Font.Color = RGB(0, 0, 0)  ' 黒文字
                End With
            End If
            If Err.Number <> 0 Then
                Err.Clear
            End If
            On Error GoTo 0
            
            ' ▼ シート「フィールド」の処理
            On Error Resume Next
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
                    ' 無限ループ防止：最大行数を制限（10000行まで）
                    If lastRow > 10000 Then
                        lastRow = 10000
                    End If
                    
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
                        ' 無限ループ防止：最大処理行数を制限
                        Dim maxProcessRows
                        maxProcessRows = lastRow - startRow + 1
                        If maxProcessRows > 10000 Then
                            maxProcessRows = 10000
                        End If
                        
                        For row = startRow To lastRow
                            ' 無限ループ防止：処理行数が上限を超えた場合は終了
                            If arrRow >= maxProcessRows Then
                                Exit For
                            End If
                            
                            ' ▼ D列以降（D列=4列目からAK列=37列目まで）に値があるかチェック
                            ' ※B列（2列目）の値の有無は無視（チェック対象外）
                            Dim hasValue, col
                            hasValue = False
                            On Error Resume Next
                            ' D列（4列目）からAK列（37列目）までチェック（B列は除外）
                            For col = 4 To 37
                                If maxCol >= col And arrRow + 1 <= maxRow Then
                                    Dim cellValue
                                    cellValue = Trim(CStr(dataArr(arrRow + 1, col)))
                                    If Err.Number = 0 And cellValue <> "" And Not IsEmpty(dataArr(arrRow + 1, col)) Then
                                        hasValue = True
                                        Exit For
                                    End If
                                    Err.Clear
                                End If
                            Next
                            On Error GoTo 0
                            
                            ' D列以降に値がない行は削除対象
                            If Not hasValue Then
                                emptyRows.Add arrRow, row
                            Else
                                ' D列以降に値がある行は処理対象
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
                                
                                ' G列に値がある場合、C列に "proto_" をセット（既存の値がない場合のみ）
                                If gValue <> "" Then
                                    On Error Resume Next
                                    If maxCol >= 3 And arrRow + 1 <= maxRow Then  ' C列は3列目
                                        cValue = Trim(CStr(dataArr(arrRow + 1, 3)))
                                        If Err.Number <> 0 Or cValue = "" Or IsEmpty(dataArr(arrRow + 1, 3)) Then
                                            dataArr(arrRow + 1, 3) = "proto_"
                                        End If
                                        Err.Clear
                                    End If
                                    On Error GoTo 0
                                    
                                    ' 「カスタム」と「標準」で分類
                                    If LCase(gValue) = "カスタム" Then
                                        customRows.Add customRows.Count, arrRow
                                    ElseIf LCase(gValue) = "標準" Then
                                        standardRows.Add standardRows.Count, arrRow
                                    Else
                                        ' その他の値も標準として扱う
                                        standardRows.Add standardRows.Count, arrRow
                                    End If
                                Else
                                    ' G列が空でもD列以降に値がある場合は標準として扱う
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
                End If
            End If
            If Err.Number <> 0 Then
                ' エラーが発生した場合でもオブジェクトを解放
                On Error Resume Next
                If Not customRows Is Nothing Then
                    Set customRows = Nothing
                End If
                If Not standardRows Is Nothing Then
                    Set standardRows = Nothing
                End If
                If Not emptyRows Is Nothing Then
                    Set emptyRows = Nothing
                End If
                dataArr = Array()
                sortedArr = Array()
                Err.Clear
            End If
            On Error GoTo 0
            
            ' シートが見つからない場合の警告
            If wsTable Is Nothing And wsField Is Nothing Then
                MsgBox "シート「テーブル」または「フィールド」が見つかりません: " & fileName, vbWarning, "警告"
            ElseIf wsTable Is Nothing Then
                MsgBox "シート「テーブル」が見つかりません: " & fileName, vbWarning, "警告"
            ElseIf wsField Is Nothing Then
                MsgBox "シート「フィールド」が見つかりません: " & fileName, vbWarning, "警告"
            End If
            
            ' 最後に「表紙」シートをアクティブにしてA1にカーソルを戻す
            On Error Resume Next
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
            
            fileProcessed = True
        Else
            MsgBox "ファイルを開けませんでした: " & fileName & vbCrLf & "エラー: " & Err.Description, vbCritical, "エラー"
            Err.Clear
            On Error GoTo 0
        End If
        
        ' ▼ エラー発生時でも確実にファイルを閉じる（クリーンアップ）
        On Error Resume Next
        If Not wb Is Nothing Then
            If Not fileProcessed Then
                ' 保存せずに閉じる
                wb.Close False
            Else
                ' 既に保存済みの場合は閉じるだけ
                wb.Close False
            End If
            Set wb = Nothing
        End If
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
        If Not customRows Is Nothing Then
            Set customRows = Nothing
        End If
        If Not standardRows Is Nothing Then
            Set standardRows = Nothing
        End If
        If Not emptyRows Is Nothing Then
            Set emptyRows = Nothing
        End If
        Err.Clear
        On Error GoTo 0
    End If
Next

' ▼ Excel終了（設定を戻してから Quit）
' エラーが発生しても確実にExcelを終了させる
On Error Resume Next
If Not excel Is Nothing Then
    excel.Calculation = -4105   ' xlCalculationAutomatic（失敗しても無視）
    excel.ScreenUpdating = True
    excel.EnableEvents = True
    excel.DisplayAlerts = True
    excel.Quit
    Set excel = Nothing
End If
Err.Clear
On Error GoTo 0

MsgBox "処理が完了しました。", vbInformation, "完了"

