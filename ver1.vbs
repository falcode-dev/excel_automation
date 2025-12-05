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
Dim primaryNameAttribute, ws2, colIndexDict2
Dim logicalNameCol, lastRow2, row, foundRow
Dim attributeMappingDict, attrFieldName, attrCellAddr, attrValue
Dim rowLogicalName

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
                            convertedValue = fieldValue
                            If IsNumeric(fieldValue) = False Then
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
                            ' セットした値を赤文字にする
                            wsTable.Cells(rowNum, colNum).Font.Color = RGB(255, 0, 0)
                        Next
                        
                        ' ▼ PrimaryNameAttributeの値を取得（シート1の3行目から）
                        primaryNameAttribute = ""
                        If colIndexDict.Exists("primarynameattribute") Then
                            On Error Resume Next
                            primaryNameAttribute = CStr(ws.Cells(4, colIndexDict("primarynameattribute")).Value2)
                            If Err.Number <> 0 Then
                                primaryNameAttribute = ""
                                Err.Clear
                            End If
                            On Error GoTo 0
                        End If
                        
                        ' ▼ シート2から該当行を検索して値をセット
                        If primaryNameAttribute <> "" And wb.Sheets.Count >= 2 Then
                            Set ws2 = wb.Sheets(2)
                            
                            ' シート2の1行目（ヘッダー行）から列インデックスを作成
                            Set colIndexDict2 = CreateObject("Scripting.Dictionary")
                            lastCol = ws2.Cells(1, ws2.Columns.Count).End(-4159).Column ' xlToLeft
                            
                            For col = 1 To lastCol
                                colName = Trim(CStr(ws2.Cells(1, col).Value2))
                                If colName <> "" Then
                                    colIndexDict2(LCase(colName)) = col
                                End If
                            Next
                            
                            ' Logical Name列の位置を取得
                            logicalNameCol = 0
                            If colIndexDict2.Exists("logical name") Then
                                logicalNameCol = colIndexDict2("logical name")
                            End If
                            
                            ' PrimaryNameAttributeと一致する行を検索
                            foundRow = 0
                            If logicalNameCol > 0 Then
                                lastRow2 = ws2.Cells(ws2.Rows.Count, logicalNameCol).End(-4162).Row ' xlUp
                                
                                For row = 2 To lastRow2
                                    On Error Resume Next
                                    rowLogicalName = Trim(CStr(ws2.Cells(row, logicalNameCol).Value2))
                                    If Err.Number <> 0 Then
                                        rowLogicalName = ""
                                        Err.Clear
                                    End If
                                    On Error GoTo 0
                                    
                                    If LCase(rowLogicalName) = LCase(primaryNameAttribute) Then
                                        foundRow = row
                                        Exit For
                                    End If
                                Next
                            End If
                            
                            ' 該当行が見つかった場合、値をセット
                            If foundRow > 0 Then
                                ' 属性マッピング辞書を作成
                                Set attributeMappingDict = CreateObject("Scripting.Dictionary")
                                attributeMappingDict.Add "Display Name", "E16"
                                attributeMappingDict.Add "Description", "E17"
                                attributeMappingDict.Add "Schema Name", "E18"
                                attributeMappingDict.Add "Logical Name", "E19"
                                attributeMappingDict.Add "Required Level", "E20"
                                attributeMappingDict.Add "Additional data", "E21"
                                
                                ' 各属性の値を取得してセット
                                For Each attrFieldName In attributeMappingDict.Keys
                                    attrCellAddr = attributeMappingDict(attrFieldName)
                                    attrValue = ""
                                    
                                    ' 列インデックスから値を取得
                                    If colIndexDict2.Exists(LCase(attrFieldName)) Then
                                        On Error Resume Next
                                        attrValue = ws2.Cells(foundRow, colIndexDict2(LCase(attrFieldName))).Value2
                                        If Err.Number <> 0 Then
                                            attrValue = ""
                                            Err.Clear
                                        End If
                                        On Error GoTo 0
                                    End If
                                    
                                    ' True/Falseを変換
                                    convertedValue = attrValue
                                    If IsNumeric(attrValue) = False Then
                                        lowerVal = LCase(Trim(CStr(attrValue)))
                                        If lowerVal = "true" Then
                                            convertedValue = ChrW(10003) ' ✓
                                        ElseIf lowerVal = "false" Or lowerVal = "" Then
                                            convertedValue = "-"
                                        End If
                                    End If
                                    
                                    ' セルアドレスを解析
                                    colNum = Asc(UCase(Left(attrCellAddr, 1))) - 64
                                    rowNum = CInt(Mid(attrCellAddr, 2))
                                    
                                    ' 値をセット（赤文字）
                                    wsTable.Cells(rowNum, colNum).Value = convertedValue
                                    wsTable.Cells(rowNum, colNum).Font.Color = RGB(255, 0, 0)
                                Next
                                
                                Set attributeMappingDict = Nothing
                            End If
                            
                            Set colIndexDict2 = Nothing
                            Set ws2 = Nothing
                        End If
                        
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
