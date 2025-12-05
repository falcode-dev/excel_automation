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
Dim rowLogicalName, pluralDisplayName
Dim maxLengthPos, maxLengthValue, afterMaxLength, i, char
Dim e13Value, wsField
Dim fieldMappingDict, outputRow, fieldValue2, outputCol
Dim additionalDataValue, targetsValue, formatValue, targetsPos, formatPos
Dim lowerFormatValue, attributeTypeValue
Dim attrTypeConverted, minValue, maxValue, optionsValue, defaultValue, targetValue, statesValue
Dim precisionPos

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
                    
                    ' シート「テーブル」「表紙」「フィールド」を取得
                    On Error Resume Next
                    Set wsTable = wbTemplate.Sheets("テーブル")
                    Set wsCover = wbTemplate.Sheets("表紙")
                    Set wsField = wbTemplate.Sheets("フィールド")
                    
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
                        
                        ' ▼ Plural Display Nameの値をE5にセット
                        pluralDisplayName = ""
                        If colIndexDict.Exists("plural display name") Then
                            On Error Resume Next
                            pluralDisplayName = ws.Cells(4, colIndexDict("plural display name")).Value2
                            If Err.Number <> 0 Then
                                pluralDisplayName = ""
                                Err.Clear
                            End If
                            On Error GoTo 0
                        End If
                        
                        If pluralDisplayName <> "" Then
                            wsTable.Cells(5, 5).Value = pluralDisplayName
                            wsTable.Cells(5, 5).Font.Color = RGB(255, 0, 0)
                        End If
                        
                        ' ▼ E13の値をチェック（空の場合は「なし」をセット）
                        e13Value = wsTable.Cells(13, 5).Value
                        If e13Value = "" Or IsEmpty(e13Value) Then
                            wsTable.Cells(13, 5).Value = "なし"
                            wsTable.Cells(13, 5).Font.Color = RGB(255, 0, 0)
                        End If
                        
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
                                    
                                    ' E21（Additional data）の場合は「Max length:」の後の数値のみを抽出
                                    If attrCellAddr = "E21" Then
                                        maxLengthValue = ""
                                        maxLengthPos = InStr(1, CStr(attrValue), "Max length:", vbTextCompare)
                                        If maxLengthPos > 0 Then
                                            afterMaxLength = Mid(CStr(attrValue), maxLengthPos + Len("Max length:"))
                                            ' 数値部分を抽出
                                            For i = 1 To Len(afterMaxLength)
                                                char = Mid(afterMaxLength, i, 1)
                                                If IsNumeric(char) Then
                                                    maxLengthValue = maxLengthValue & char
                                                ElseIf maxLengthValue <> "" Then
                                                    Exit For
                                                End If
                                            Next
                                        End If
                                        convertedValue = maxLengthValue
                                    Else
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
                        End If
                        
                        ' ▼ シート2の2行目以降のデータをシート「フィールド」に出力
                        If wb.Sheets.Count >= 2 Then
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
                            
                            ' フィールドマッピング辞書を作成（フィールド名 → 列番号）
                            Set fieldMappingDict = CreateObject("Scripting.Dictionary")
                            fieldMappingDict.Add "Schema Name", 4    ' D列
                            fieldMappingDict.Add "Display Name", 5   ' E列
                            fieldMappingDict.Add "Custom Attribute", 7 ' G列
                            fieldMappingDict.Add "Attribute Type", 10  ' J列
                            fieldMappingDict.Add "Type", 11            ' K列
                            fieldMappingDict.Add "Required Level", 12  ' L列
                            fieldMappingDict.Add "Description", 31     ' AE列
                            fieldMappingDict.Add "Audit Enabled", 24  ' X列
                            fieldMappingDict.Add "Secured", 25         ' Y列
                            fieldMappingDict.Add "ValidFor AdvancedFind", 28 ' AB列
                            
                            ' シート2の最終行を取得
                            lastRow2 = ws2.Cells(ws2.Rows.Count, 1).End(-4162).Row ' xlUp
                            
                            ' 2行目以降のデータを処理
                            outputRow = 7 ' 出力開始行
                            
                            For row = 2 To lastRow2
                                ' 各フィールドの値を取得してセット
                                For Each fieldName In fieldMappingDict.Keys
                                    outputCol = fieldMappingDict(fieldName)
                                    fieldValue2 = ""
                                    
                                    ' 列インデックスから値を取得
                                    If colIndexDict2.Exists(LCase(fieldName)) Then
                                        On Error Resume Next
                                        fieldValue2 = ws2.Cells(row, colIndexDict2(LCase(fieldName))).Value2
                                        If Err.Number <> 0 Then
                                            fieldValue2 = ""
                                            Err.Clear
                                        End If
                                        On Error GoTo 0
                                    End If
                                    
                                    ' フィールドごとの変換処理（空の値は空のまま）
                                    convertedValue = fieldValue2
                                    
                                    ' 空の値の場合は変換せずに空のまま
                                    If fieldValue2 = "" Or IsEmpty(fieldValue2) Then
                                        convertedValue = ""
                                    ElseIf IsNumeric(fieldValue2) = False Then
                                        lowerVal = LCase(Trim(CStr(fieldValue2)))
                                        
                                        ' Custom Attributeの変換（True → カスタム、False → 標準）
                                        If fieldName = "Custom Attribute" Then
                                            If lowerVal = "true" Then
                                                convertedValue = "カスタム"
                                            ElseIf lowerVal = "false" Then
                                                convertedValue = "標準"
                                            End If
                                        ' Typeの変換（Simple → シンプル、Calculated → 計算、Rollup → ロールアップ）
                                        ElseIf fieldName = "Type" Then
                                            Select Case lowerVal
                                                Case "simple"
                                                    convertedValue = "シンプル"
                                                Case "calculated"
                                                    convertedValue = "計算"
                                                Case "rollup"
                                                    convertedValue = "ロールアップ"
                                                Case Else
                                                    convertedValue = fieldValue2
                                            End Select
                                        ' Required Levelの変換（None → 任意、ApplicationRequired → 必須項目、Recommended → 推奨項目）
                                        ElseIf fieldName = "Required Level" Then
                                            Select Case lowerVal
                                                Case "none"
                                                    convertedValue = "任意"
                                                Case "applicationrequired"
                                                    convertedValue = "システム要求"
                                                Case "systemrequired"
                                                    convertedValue = "必須項目"
                                                Case "recommended"
                                                    convertedValue = "推奨項目"
                                                Case Else
                                                    convertedValue = fieldValue2
                                            End Select
                                        ' その他のTrue/FalseはTRUE/FALSEに変換（シート「フィールド」用）
                                        Else
                                            If lowerVal = "true" Then
                                                convertedValue = "TRUE"
                                            ElseIf lowerVal = "false" Then
                                                convertedValue = "FALSE"
                                            End If
                                        End If
                                    End If
                                    
                                    ' 値をセット（赤文字）
                                    wsField.Cells(outputRow, outputCol).Value = convertedValue
                                    wsField.Cells(outputRow, outputCol).Font.Color = RGB(255, 0, 0)
                                Next
                                
                                ' ▼ Additional dataの処理
                                additionalDataValue = ""
                                
                                ' Additional dataの値を取得
                                If colIndexDict2.Exists("additional data") Then
                                    On Error Resume Next
                                    additionalDataValue = CStr(ws2.Cells(row, colIndexDict2("additional data")).Value2)
                                    If Err.Number <> 0 Then
                                        additionalDataValue = ""
                                        Err.Clear
                                    End If
                                    On Error GoTo 0
                                    
                                    ' Precision:以降を削除
                                    Dim precisionPos
                                    precisionPos = InStr(1, additionalDataValue, "Precision:", vbTextCompare)
                                    If precisionPos > 0 Then
                                        additionalDataValue = Left(additionalDataValue, precisionPos - 1)
                                    End If
                                End If
                                
                                ' targets:の処理（V7 = 22列目）
                                If additionalDataValue <> "" Then
                                    targetsPos = InStr(1, additionalDataValue, "targets:", vbTextCompare)
                                    If targetsPos > 0 Then
                                        targetsValue = Mid(additionalDataValue, targetsPos + Len("targets:"))
                                        ' 改行やスペースを取り除く
                                        targetsValue = Replace(targetsValue, vbCrLf, "")
                                        targetsValue = Replace(targetsValue, vbLf, "")
                                        targetsValue = Replace(targetsValue, vbCr, "")
                                        targetsValue = Replace(targetsValue, " ", "")
                                        targetsValue = Trim(targetsValue)
                                        
                                        ' V7（22列目）にセット
                                        wsField.Cells(outputRow, 22).Value = targetsValue
                                        wsField.Cells(outputRow, 22).Font.Color = RGB(255, 0, 0)
                                    End If
                                    
                                    ' Format:の処理（J8 = 10列目、8行目にセット）
                                    formatPos = InStr(1, additionalDataValue, "Format:", vbTextCompare)
                                    If formatPos > 0 Then
                                        formatValue = Mid(additionalDataValue, formatPos + Len("Format:"))
                                        ' 改行やスペースを取り除く
                                        formatValue = Replace(formatValue, vbCrLf, "")
                                        formatValue = Replace(formatValue, vbLf, "")
                                        formatValue = Replace(formatValue, vbCr, "")
                                        formatValue = Trim(formatValue)
                                        
                                        ' DateAndTime/DateOnlyの変換
                                        lowerFormatValue = LCase(formatValue)
                                        If InStr(lowerFormatValue, "dateandtime") > 0 Then
                                            formatValue = "日付と時刻 - 日時"
                                        ElseIf InStr(lowerFormatValue, "dateonly") > 0 Then
                                            formatValue = "日付と時刻 - 日付のみ"
                                        End If
                                        
                                        ' J8（10列目、8行目）にセット
                                        wsField.Cells(8, 10).Value = formatValue
                                        wsField.Cells(8, 10).Font.Color = RGB(255, 0, 0)
                                    End If
                                    
                                    ' ▼ Attribute Typeの変換とAdditional dataの処理
                                    attrTypeConverted = ""
                                    minValue = ""
                                    maxValue = ""
                                    optionsValue = ""
                                    defaultValue = ""
                                    targetValue = ""
                                    statesValue = ""
                                    
                                    ' Attribute Typeの値を取得
                                    If colIndexDict2.Exists("attribute type") Then
                                        On Error Resume Next
                                        attributeTypeValue = CStr(ws2.Cells(row, colIndexDict2("attribute type")).Value2)
                                        If Err.Number <> 0 Then
                                            attributeTypeValue = ""
                                            Err.Clear
                                        End If
                                        On Error GoTo 0
                                        
                                        ' Attribute Typeの変換
                                        lowerVal = LCase(Trim(attributeTypeValue))
                                        Select Case lowerVal
                                            Case "bigint"
                                                attrTypeConverted = "数値 - 整数(Int)"
                                                ' Minimum value/Maximum valueを抽出
                                                minValue = ExtractValueFromAdditionalData(additionalDataValue, "Minimum value:")
                                                maxValue = ExtractValueFromAdditionalData(additionalDataValue, "Maximum value:")
                                            Case "choice"
                                                attrTypeConverted = "選択肢"
                                                optionsValue = ExtractValueFromAdditionalData(additionalDataValue, "Options:")
                                                defaultValue = ExtractValueFromAdditionalData(additionalDataValue, "Default:")
                                                If LCase(Trim(defaultValue)) = "n/a" Then
                                                    defaultValue = "なし"
                                                End If
                                            Case "choices"
                                                attrTypeConverted = "選択肢(複数)"
                                                optionsValue = ExtractValueFromAdditionalData(additionalDataValue, "Options:")
                                                defaultValue = ExtractValueFromAdditionalData(additionalDataValue, "Default:")
                                                If LCase(Trim(defaultValue)) = "n/a" Then
                                                    defaultValue = "なし"
                                                End If
                                            Case "currency"
                                                attrTypeConverted = "通貨"
                                                minValue = ExtractValueFromAdditionalData(additionalDataValue, "Minimum value:")
                                                maxValue = ExtractValueFromAdditionalData(additionalDataValue, "Maximum value:")
                                            Case "decimal"
                                                attrTypeConverted = "数値 - 少数(10進数)"
                                                minValue = ExtractValueFromAdditionalData(additionalDataValue, "Minimum value:")
                                                maxValue = ExtractValueFromAdditionalData(additionalDataValue, "Maximum value:")
                                            Case "double"
                                                attrTypeConverted = "数値 - 浮動小数点数"
                                                minValue = ExtractValueFromAdditionalData(additionalDataValue, "Minimum value:")
                                                maxValue = ExtractValueFromAdditionalData(additionalDataValue, "Maximum value:")
                                            Case "multiline text"
                                                attrTypeConverted = "複数行テキスト - プレーン"
                                            Case "owner"
                                                attrTypeConverted = "所有者"
                                                targetValue = ExtractValueFromAdditionalData(additionalDataValue, "Target:")
                                            Case "state"
                                                attrTypeConverted = "状態"
                                                statesValue = ExtractValueFromAdditionalData(additionalDataValue, "States:")
                                            Case "status"
                                                attrTypeConverted = "ステータス"
                                                statesValue = ExtractValueFromAdditionalData(additionalDataValue, "States:")
                                            Case "text"
                                                attrTypeConverted = "1行テキスト - プレーン"
                                            Case "two options"
                                                attrTypeConverted = "はい/いいえ"
                                                optionsValue = ExtractValueFromAdditionalData(additionalDataValue, "Options:")
                                                defaultValue = ExtractValueFromAdditionalData(additionalDataValue, "Default Value:")
                                            Case "uniqueidentifier":
                                                attrTypeConverted = "一意識別子"
                                            Case "whole number"
                                                attrTypeConverted = "数値 - 整数(Int)"
                                                minValue = ExtractValueFromAdditionalData(additionalDataValue, "Minimum value")
                                                maxValue = ExtractValueFromAdditionalData(additionalDataValue, "Maximum value")
                                            Case Else
                                                ' Lookupの場合は検索に変換（既存の処理）
                                                If lowerVal = "lookup" Then
                                                    attrTypeConverted = "検索"
                                                Else
                                                    attrTypeConverted = attributeTypeValue
                                                End If
                                        End Select
                                        
                                        ' Attribute TypeをJ列（10列目）にセット
                                        If attrTypeConverted <> "" Then
                                            wsField.Cells(outputRow, 10).Value = attrTypeConverted
                                            wsField.Cells(outputRow, 10).Font.Color = RGB(255, 0, 0)
                                        End If
                                        
                                        ' Additional dataの値を各セルにセット
                                        ' Minimum value → P7（16列目）
                                        If minValue <> "" Then
                                            wsField.Cells(outputRow, 16).Value = minValue
                                            wsField.Cells(outputRow, 16).Font.Color = RGB(255, 0, 0)
                                        End If
                                        
                                        ' Maximum value → O7（15列目）
                                        If maxValue <> "" Then
                                            wsField.Cells(outputRow, 15).Value = maxValue
                                            wsField.Cells(outputRow, 15).Font.Color = RGB(255, 0, 0)
                                        End If
                                        
                                        ' Options: → T7（20列目）
                                        If optionsValue <> "" Then
                                            wsField.Cells(outputRow, 20).Value = optionsValue
                                            wsField.Cells(outputRow, 20).Font.Color = RGB(255, 0, 0)
                                        End If
                                        
                                        ' Default: / Default Value: → U7（21列目）
                                        If defaultValue <> "" Then
                                            wsField.Cells(outputRow, 21).Value = defaultValue
                                            wsField.Cells(outputRow, 21).Font.Color = RGB(255, 0, 0)
                                        End If
                                        
                                        ' Target: → V7（22列目）
                                        If targetValue <> "" Then
                                            wsField.Cells(outputRow, 22).Value = targetValue
                                            wsField.Cells(outputRow, 22).Font.Color = RGB(255, 0, 0)
                                        End If
                                        
                                        ' States: → T7（20列目）
                                        If statesValue <> "" Then
                                            wsField.Cells(outputRow, 20).Value = statesValue
                                            wsField.Cells(outputRow, 20).Font.Color = RGB(255, 0, 0)
                                        End If
                                        
                                        ' Multiline Text/Textの場合はAdditional data全体をAK7（37列目）にセット
                                        If lowerVal = "multiline text" Or lowerVal = "text" Then
                                            If additionalDataValue <> "" Then
                                                wsField.Cells(outputRow, 37).Value = additionalDataValue
                                                wsField.Cells(outputRow, 37).Font.Color = RGB(255, 0, 0)
                                            End If
                                        End If
                                    End If
                                End If
                                
                                
                                outputRow = outputRow + 1
                            Next
                            
                            Set fieldMappingDict = Nothing
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
                        Set wsField = Nothing
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

'────────────────────────────────────────
'  関数：Additional dataから値を抽出
'────────────────────────────────────────
Function ExtractValueFromAdditionalData(additionalData, keyword)
    Dim pos, valueStart, valueEnd, nextKeywordPos
    Dim result, tempValue
    Dim keywords
    Dim i, keywordPos, minPos
    
    result = ""
    pos = InStr(1, additionalData, keyword, vbTextCompare)
    
    If pos > 0 Then
        valueStart = pos + Len(keyword)
        
        ' 次のキーワードの位置を探す
        keywords = Array("Minimum value:", "Maximum value:", "Options:", "Default:", "Default Value:", "Target:", "States:", "Format:", "targets:")
        minPos = 0
        
        For i = 0 To UBound(keywords)
            keywordPos = InStr(valueStart, additionalData, keywords(i), vbTextCompare)
            If keywordPos > 0 Then
                If minPos = 0 Or keywordPos < minPos Then
                    minPos = keywordPos
                End If
            End If
        Next
        
        If minPos > 0 Then
            tempValue = Mid(additionalData, valueStart, minPos - valueStart)
        Else
            tempValue = Mid(additionalData, valueStart)
        End If
        
        ' 改行やスペースを取り除く
        tempValue = Replace(tempValue, vbCrLf, "")
        tempValue = Replace(tempValue, vbLf, "")
        tempValue = Replace(tempValue, vbCr, "")
        tempValue = Replace(tempValue, " ", "")
        result = Trim(tempValue)
    End If
    
    ExtractValueFromAdditionalData = result
End Function
