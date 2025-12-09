Option Explicit

'────────────────────────────────────────
'  処理：ドラッグ&ドロップでフォルダを受け取り、
'  フォルダ内のExcelファイルを順次処理し、
'  ・C7以降のセルの値を空にする
'  ・O7とP7以降の表示形式を「数値」にする
'  ・シート「テーブル」「フィールド」のセルから「msdyn_」を「tel_wo」に置換
'  ・シート「テーブル」のE25を空にする
'  ・シート「フォーム」と「ビュー」を削除
'  ・同じ階層にある template.xlsx のシート「フォーム」「ビュー」を
'    シート「フィールド」の後ろにコピーする
'  ※エラーハンドリングと処理のベースは
'    ExcelFormat.vbsを参考にしています
'────────────────────────────────────────

Dim fso, excel, wb, wsTable, wsField, ws, wsForm, wsView
Dim folderPath, folder, file
Dim fileName, filePath, fileExt
Dim lastRow, lastCol, lastRowP
Dim templatePath, templateWb
Dim fieldSheetIndex
Dim cellValue, replacedValue
Dim i, j
Dim templateForm, templateView

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

' ▼ template.xlsx のパスを取得（同じ階層）
templatePath = fso.GetParentFolderName(folderPath) & "\template.xlsx"

' ▼ template.xlsx の存在チェック
If Not fso.FileExists(templatePath) Then
    MsgBox "template.xlsx が見つかりません: " & templatePath, vbCritical, "エラー"
    WScript.Quit
End If

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
    
    ' Excelファイルの拡張子をチェック（template.xlsxは除外）
    If (fileExt = "xlsx" Or fileExt = "xls" Or fileExt = "xlsm" Or fileExt = "xlsb") And LCase(fileName) <> "template.xlsx" Then
        filePath = file.Path
        
        ' 前のループで作成されたオブジェクトを解放（メモリリーク防止）
        On Error Resume Next
        If Not wsTable Is Nothing Then
            Set wsTable = Nothing
        End If
        If Not wsField Is Nothing Then
            Set wsField = Nothing
        End If
        If Not wsForm Is Nothing Then
            Set wsForm = Nothing
        End If
        If Not wsView Is Nothing Then
            Set wsView = Nothing
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
            
            ' シート「テーブル」「フィールド」「フォーム」「ビュー」を取得
            On Error Resume Next
            Set wsTable = Nothing
            Set wsField = Nothing
            Set wsForm = Nothing
            Set wsView = Nothing
            
            ' シート名で検索
            For Each ws In wb.Sheets
                If ws.Name = "テーブル" Then
                    Set wsTable = ws
                ElseIf ws.Name = "フィールド" Then
                    Set wsField = ws
                ElseIf ws.Name = "フォーム" Then
                    Set wsForm = ws
                ElseIf ws.Name = "ビュー" Then
                    Set wsView = ws
                End If
            Next
            
            If Err.Number <> 0 Then
                Err.Clear
            End If
            On Error GoTo 0
            
            ' ▼ シート「フィールド」の処理
            If Not wsField Is Nothing Then
                ' ▼ C7以降のセルの値を空にする
                ' 最終行を取得（C列で判定）
                On Error Resume Next
                lastRow = wsField.Cells(wsField.Rows.Count, 3).End(-4162).Row ' xlUp (C列=3列目)
                If Err.Number <> 0 Or lastRow < 7 Then
                    lastRow = 7
                    Err.Clear
                End If
                On Error GoTo 0
                
                ' C7以降をクリア
                If lastRow >= 7 Then
                    On Error Resume Next
                    wsField.Range(wsField.Cells(7, 3), wsField.Cells(lastRow, 3)).ClearContents
                    If Err.Number <> 0 Then
                        Err.Clear
                    End If
                    On Error GoTo 0
                End If
                
                ' ▼ O7とP7以降の表示形式を「数値」にする（マイナス値も対応）
                ' O列（15列目）とP列（16列目）の最終行を取得
                On Error Resume Next
                lastRow = wsField.Cells(wsField.Rows.Count, 15).End(-4162).Row ' xlUp (O列=15列目)
                lastRowP = wsField.Cells(wsField.Rows.Count, 16).End(-4162).Row ' xlUp (P列=16列目)
                If Err.Number <> 0 Or lastRow < 7 Then
                    lastRow = 7
                    Err.Clear
                End If
                If lastRowP < 7 Then
                    lastRowP = 7
                End If
                ' より大きい方の最終行を使用
                If lastRowP > lastRow Then
                    lastRow = lastRowP
                End If
                On Error GoTo 0
                
                ' O7とP7以降の表示形式を数値に設定（マイナス値も表示可能な形式）
                If lastRow >= 7 Then
                    On Error Resume Next
                    ' O列（15列目）の表示形式を数値に設定（"0"形式でマイナス値も表示可能）
                    wsField.Range(wsField.Cells(7, 15), wsField.Cells(lastRow, 15)).NumberFormat = "0"
                    ' P列（16列目）の表示形式を数値に設定（"0"形式でマイナス値も表示可能）
                    wsField.Range(wsField.Cells(7, 16), wsField.Cells(lastRow, 16)).NumberFormat = "0"
                    If Err.Number <> 0 Then
                        Err.Clear
                    End If
                    On Error GoTo 0
                End If
            End If
            
            ' ▼ シート「テーブル」と「フィールド」のセルから「msdyn_」を「tel_wo」に置換
            ' シート「テーブル」の処理
            If Not wsTable Is Nothing Then
                On Error Resume Next
                ' 使用されている範囲を取得
                lastRow = wsTable.UsedRange.Rows.Count
                lastCol = wsTable.UsedRange.Columns.Count
                If Err.Number <> 0 Or lastRow = 0 Or lastCol = 0 Then
                    lastRow = 1
                    lastCol = 1
                    Err.Clear
                End If
                On Error GoTo 0
                
                ' すべてのセルをチェックして置換
                If lastRow > 0 And lastCol > 0 Then
                    For i = 1 To lastRow
                        For j = 1 To lastCol
                            On Error Resume Next
                            cellValue = CStr(wsTable.Cells(i, j).Value)
                            If Err.Number = 0 And Not IsEmpty(wsTable.Cells(i, j).Value) Then
                                If InStr(1, cellValue, "msdyn_", vbTextCompare) > 0 Then
                                    replacedValue = Replace(cellValue, "msdyn_", "tel_wo", 1, -1, vbTextCompare)
                                    wsTable.Cells(i, j).Value = replacedValue
                                End If
                            End If
                            Err.Clear
                            On Error GoTo 0
                        Next
                    Next
                End If
            End If
            
            ' シート「フィールド」の処理
            If Not wsField Is Nothing Then
                On Error Resume Next
                ' 使用されている範囲を取得
                lastRow = wsField.UsedRange.Rows.Count
                lastCol = wsField.UsedRange.Columns.Count
                If Err.Number <> 0 Or lastRow = 0 Or lastCol = 0 Then
                    lastRow = 1
                    lastCol = 1
                    Err.Clear
                End If
                On Error GoTo 0
                
                ' すべてのセルをチェックして置換
                If lastRow > 0 And lastCol > 0 Then
                    For i = 1 To lastRow
                        For j = 1 To lastCol
                            On Error Resume Next
                            cellValue = CStr(wsField.Cells(i, j).Value)
                            If Err.Number = 0 And Not IsEmpty(wsField.Cells(i, j).Value) Then
                                If InStr(1, cellValue, "msdyn_", vbTextCompare) > 0 Then
                                    replacedValue = Replace(cellValue, "msdyn_", "tel_wo", 1, -1, vbTextCompare)
                                    wsField.Cells(i, j).Value = replacedValue
                                End If
                            End If
                            Err.Clear
                            On Error GoTo 0
                        Next
                    Next
                End If
            End If
            
            ' ▼ シート「テーブル」のE25を空にする
            If Not wsTable Is Nothing Then
                On Error Resume Next
                wsTable.Cells(25, 5).Value = ""  ' E25 = 5列目、25行目
                If Err.Number <> 0 Then
                    Err.Clear
                End If
                On Error GoTo 0
            End If
            
            ' ▼ シート「フォーム」と「ビュー」を削除
            On Error Resume Next
            If Not wsForm Is Nothing Then
                excel.DisplayAlerts = False
                wsForm.Delete
                excel.DisplayAlerts = False
            End If
            If Not wsView Is Nothing Then
                excel.DisplayAlerts = False
                wsView.Delete
                excel.DisplayAlerts = False
            End If
            If Err.Number <> 0 Then
                Err.Clear
            End If
            On Error GoTo 0
            
            ' ▼ template.xlsx のシート「フォーム」「ビュー」をシート「フィールド」の後ろにコピー
            If Not wsField Is Nothing Then
                ' 「フィールド」シートのインデックスを取得
                On Error Resume Next
                fieldSheetIndex = wsField.Index
                If Err.Number <> 0 Then
                    fieldSheetIndex = 1
                    Err.Clear
                End If
                On Error GoTo 0
                
                ' template.xlsx を開く
                On Error Resume Next
                Set templateWb = excel.Workbooks.Open(templatePath, 0, False)
                If Err.Number = 0 Then
                    On Error GoTo 0
                    
                    ' template.xlsx から「フォーム」「ビュー」シートを探してコピー
                    Set templateForm = Nothing
                    Set templateView = Nothing
                    
                    For Each ws In templateWb.Sheets
                        If ws.Name = "フォーム" Then
                            Set templateForm = ws
                        ElseIf ws.Name = "ビュー" Then
                            Set templateView = ws
                        End If
                    Next
                    
                    ' 「フォーム」シートをコピー
                    If Not templateForm Is Nothing Then
                        On Error Resume Next
                        templateForm.Copy , wb.Sheets(fieldSheetIndex)
                        If Err.Number <> 0 Then
                            Err.Clear
                        End If
                        On Error GoTo 0
                    End If
                    
                    ' 「ビュー」シートをコピー
                    If Not templateView Is Nothing Then
                        On Error Resume Next
                        templateView.Copy , wb.Sheets(fieldSheetIndex)
                        If Err.Number <> 0 Then
                            Err.Clear
                        End If
                        On Error GoTo 0
                    End If
                    
                    ' template.xlsx を閉じる（保存しない）
                    templateWb.Close False
                    Set templateWb = Nothing
                Else
                    MsgBox "template.xlsx を開けませんでした: " & fileName, vbWarning, "警告"
                    Err.Clear
                    On Error GoTo 0
                End If
            End If
            
            ' 最後に「フィールド」シートをアクティブにしてA1にカーソルをセット
            If Not wsField Is Nothing Then
                wsField.Activate
                wsField.Range("A1").Select
            ElseIf Not wsTable Is Nothing Then
                wsTable.Activate
                wsTable.Range("A1").Select
            End If
            
            ' ▼ 「保存しますか？」ダイアログを防ぐため、保存前に計算を実行
            On Error Resume Next
            excel.Calculation = -4105   ' xlCalculationAutomatic
            wb.Calculate   ' ブック全体を計算
            If Err.Number <> 0 Then
                excel.Calculation = -4135   ' xlCalculationManual
                Err.Clear
            End If
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
            Set wsForm = Nothing
            Set wsView = Nothing
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

' オブジェクトを解放
Set fso = Nothing
Set folder = Nothing

MsgBox "処理が完了しました。", vbInformation, "完了"

