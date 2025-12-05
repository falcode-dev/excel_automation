Option Explicit

'────────────────────────────────────────
'  処理：ドラッグ&ドロップでフォルダを受け取り、
'  フォルダ内のExcelファイルからDisplayNameを取得し、
'  template.xlsxをベースに出力ファイルを作成
'────────────────────────────────────────

Dim fso, excel, wb, ws, wbTemplate, wsTable, wsCover
Dim folderPath, folder, file
Dim fileName, filePath, fileExt
Dim templatePath, outputFolderPath, outputFilePath
Dim displayName, displayNameCol
Dim lastCol, col, colName
Dim outputFileName
Dim mappingDict, colIndexDict
Dim fieldName, cellAddr, fieldValue
Dim rowNum, colNum
Dim convertedValue, lowerVal

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

' ▼ template.xlsxのパスを取得（同じ階層）
templatePath = fso.BuildPath(fso.GetParentFolderName(folderPath), "template.xlsx")

' ▼ template.xlsxの存在チェック
If Not fso.FileExists(templatePath) Then
    MsgBox "template.xlsxが見つかりません: " & templatePath, vbCritical, "エラー"
    WScript.Quit
End If

' ▼ 出力フォルダのパスを取得（同じ階層の「20_作成済定義書」）
outputFolderPath = fso.BuildPath(fso.GetParentFolderName(folderPath), "20_作成済定義書")

' ▼ 出力フォルダが存在しない場合は作成
If Not fso.FolderExists(outputFolderPath) Then
    fso.CreateFolder outputFolderPath
End If

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
        
        If Err.Number = 0 Then
            On Error GoTo 0
            
            ' 1シート目を取得
            Set ws = wb.Sheets(1)
            
            ' マッピング辞書を作成（フィールド名 → セルアドレス）
            Set mappingDict = CreateObject("Scripting.Dictionary")
            mappingDict.Add "Schema Name", "E9"
            mappingDict.Add "Logical Name", "E10"
            mappingDict.Add "Ownership Type", "E12"
            mappingDict.Add "Change TrackingEnabled", "E23"
            mappingDict.Add "Description", "E7"
            mappingDict.Add "DisplayCollectionName", "E6"
            mappingDict.Add "EntityColor", "E14"
            mappingDict.Add "EntityHelpUrl", "E25"
            mappingDict.Add "EntityHelpUrlEnabled", "E24"
            mappingDict.Add "HasEmailAddresses", "E35"
            mappingDict.Add "HasFeedback", "E37"
            mappingDict.Add "HasNotes", "E8"
            mappingDict.Add "IsActivity", "E30"
            mappingDict.Add "IsAuditEnabled", "E26"
            mappingDict.Add "IsAvailableOffline", "E39"
            mappingDict.Add "IsConnectionsEnabled", "E34"
            mappingDict.Add "IsDocumentManagementEnabled", "E33"
            mappingDict.Add "IsDuplicateDetection Enabled", "E22"
            mappingDict.Add "IsMailMergeEnabled", "E31"
            mappingDict.Add "IsQuickCreateEnabled", "E27"
            mappingDict.Add "IsRetentionEnabled", "E28"
            mappingDict.Add "IsSLAEnabled", "E32"
            mappingDict.Add "IsValidForAdvancedFind", "E38"
            mappingDict.Add "IsValidForQueue", "E40"
            mappingDict.Add "PrimarylmageAttribute", "E15"
            mappingDict.Add "TableType", "E11"
            
            ' 3行目（ヘッダー行）から列インデックスを作成
            Set colIndexDict = CreateObject("Scripting.Dictionary")
            lastCol = ws.Cells(3, ws.Columns.Count).End(-4159).Column ' xlToLeft
            
            For col = 1 To lastCol
                colName = Trim(CStr(ws.Cells(3, col).Value2))
                If colName <> "" Then
                    colIndexDict(LCase(colName)) = col
                End If
            Next
            
            ' DisplayNameを取得（ファイル名生成用）
            displayName = ""
            If colIndexDict.Exists("displayname") Then
                On Error Resume Next
                displayName = CStr(ws.Cells(4, colIndexDict("displayname")).Value2)
                If Err.Number <> 0 Then
                    displayName = ""
                    Err.Clear
                End If
                On Error GoTo 0
            End If
            
            ' DisplayNameが見つかった場合のみ処理
            If displayName <> "" Then
                ' template.xlsxを開く
                On Error Resume Next
                Set wbTemplate = excel.Workbooks.Open(templatePath, 0, False)
                
                If Err.Number = 0 Then
                    On Error GoTo 0
                    
                    ' シート「テーブル」と「表紙」を取得
                    On Error Resume Next
                    Set wsTable = wbTemplate.Sheets("テーブル")
                    Set wsCover = wbTemplate.Sheets("表紙")
                    
                    If Err.Number = 0 Then
                        On Error GoTo 0
                        
                        ' マッピングに従って値をセット
                        For Each fieldName In mappingDict.Keys
                            cellAddr = mappingDict(fieldName)
                            fieldValue = ""
                            
                            ' 列インデックスから値を取得
                            If colIndexDict.Exists(LCase(fieldName)) Then
                                On Error Resume Next
                                fieldValue = ws.Cells(4, colIndexDict(LCase(fieldName))).Value2
                                If Err.Number <> 0 Then
                                    fieldValue = ""
                                    Err.Clear
                                End If
                                On Error GoTo 0
                            End If
                            
                            ' True/Falseを変換（True → ✓、False → -）
                            Dim convertedValue
                            convertedValue = fieldValue
                            If IsNumeric(fieldValue) = False Then
                                Dim lowerVal
                                lowerVal = LCase(Trim(CStr(fieldValue)))
                                If lowerVal = "true" Then
                                    convertedValue = ChrW(10003) ' ✓
                                ElseIf lowerVal = "false" Or lowerVal = "" Then
                                    convertedValue = "-"
                                End If
                            End If
                            
                            ' セルアドレスを解析（例：E5 → 行5、列5）
                            colNum = Asc(UCase(Left(cellAddr, 1))) - 64 ' A=1, B=2, ..., E=5
                            rowNum = CInt(Mid(cellAddr, 2))
                            
                            ' 値をセット
                            wsTable.Cells(rowNum, colNum).Value = convertedValue
                        Next
                        
                        ' シート「表紙」のB7に「エンティティ定義書_ID_<DisplayNameの値>_v0.1」をセット
                        wsCover.Cells(7, 2).Value = "エンティティ定義書_ID_" & displayName & "_v0.1"
                        
                        ' 出力ファイル名を生成
                        outputFileName = "エンティティ定義書_ID_" & displayName & "_v0.1.xlsx"
                        outputFilePath = fso.BuildPath(outputFolderPath, outputFileName)
                        
                        ' 既存ファイルがある場合は削除
                        If fso.FileExists(outputFilePath) Then
                            fso.DeleteFile outputFilePath, True
                        End If
                        
                        ' ファイルを保存
                        wbTemplate.SaveAs outputFilePath
                        
                        ' template.xlsxを閉じる（変更を保存しない）
                        wbTemplate.Close False
                        Set wbTemplate = Nothing
                        Set wsTable = Nothing
                        Set wsCover = Nothing
                    Else
                        MsgBox "シート「テーブル」または「表紙」が見つかりません: " & fileName, vbCritical, "エラー"
                        Err.Clear
                        On Error GoTo 0
                        If Not wbTemplate Is Nothing Then
                            wbTemplate.Close False
                            Set wbTemplate = Nothing
                        End If
                    End If
                Else
                    MsgBox "template.xlsxを開けませんでした: " & Err.Description, vbCritical, "エラー"
                    Err.Clear
                    On Error GoTo 0
                End If
            Else
                MsgBox "DisplayNameが見つかりませんでした: " & fileName, vbWarning, "警告"
            End If
            
            ' 元のファイルを閉じる
            wb.Close False
            Set wb = Nothing
            Set ws = Nothing
        Else
            MsgBox "ファイルを開けませんでした: " & fileName & vbCrLf & "エラー: " & Err.Description, vbCritical, "エラー"
            Err.Clear
            On Error GoTo 0
        End If
    End If
Next

' ▼ Excel終了
excel.Quit
Set excel = Nothing

MsgBox "処理が完了しました。", vbInformation, "完了"
