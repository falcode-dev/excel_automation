Option Explicit

'────────────────────────────────────────
'  メイン処理：メタデータ結合マクロ
'────────────────────────────────────────
Public Sub メイン処理_メタデータ結合()

    Dim wbThis As Workbook
    Dim folderBase As String
    Dim folderBrowser As String
    Dim folderDocument As String
    Dim folderAttribute As String
    Dim folderOutput As String
    
    Dim fileName As String
    Dim browserPath As String
    Dim documentPath As String
    Dim attributePath As String
    Dim outputPath As String
    
    Dim wbOut As Workbook
    Dim wbBrowser As Workbook
    Dim wbDocument As Workbook
    Dim wbAttribute As Workbook
    
    Dim logicalName As String
    
    On Error GoTo ERR_HANDLER
    
    Set wbThis = ThisWorkbook
    folderBase = wbThis.Path & "\"
    
    '▼ 必要なフォルダ
    folderBrowser = folderBase & "00_preparation_work\10_metadata_browser\"
    folderDocument = folderBase & "00_preparation_work\20_metadata_document_generator\"
    folderAttribute = folderBase & "00_preparation_work\30_attribute\"
    folderOutput = folderBase & "00_preparation_work\40_generate_file\"
    
    '▼ フォルダ存在チェック
    If Dir(folderBrowser, vbDirectory) = "" Then Err.Raise 100, , "10_metadata_browser フォルダがありません。"
    If Dir(folderDocument, vbDirectory) = "" Then Err.Raise 101, , "20_metadata_document_generator フォルダがありません。"
    If Dir(folderAttribute, vbDirectory) = "" Then Err.Raise 102, , "30_attribute フォルダがありません。"
    If Dir(folderOutput, vbDirectory) = "" Then Err.Raise 103, , "40_generate_file フォルダがありません。"
    
    '▼ browser フォルダの全Excelを処理
    fileName = Dir(folderBrowser & "*.xlsx")
    If fileName = "" Then
        fileName = Dir(folderBrowser & "*.xls")
    End If
    If fileName = "" Then Err.Raise 104, , "10_metadata_browser に処理対象ファイルがありません。"
    
    '▼ ファイル処理ループ
    Do While fileName <> ""
        
        browserPath = folderBrowser & fileName
        documentPath = folderDocument & fileName
        attributePath = folderAttribute & fileName
        
        '▼ ファイル存在チェック
        If Dir(browserPath) = "" Then
            fileName = Dir()
            GoTo NEXT_FILE
        End If
        
        '▼ 新しいワークブックを作成
        Set wbOut = Workbooks.Add
        Application.DisplayAlerts = False
        
        '▼ 必要なシートを作成
        Dim wsTable As Worksheet
        Dim wsForm As Worksheet
        Set wsTable = wbOut.Sheets.Add
        wsTable.Name = "テーブル"
        Set wsForm = wbOut.Sheets.Add
        wsForm.Name = "フォーム"
        
        '▼ デフォルトシートを削除（「テーブル」と「フォーム」以外）
        '▼ 後ろから削除することでインデックスのずれを防ぐ
        Dim i As Long
        Dim sheetCount As Long
        sheetCount = wbOut.Sheets.Count
        
        For i = sheetCount To 1 Step -1
            Dim ws As Worksheet
            Set ws = wbOut.Sheets(i)
            If ws.Name <> "テーブル" And ws.Name <> "フォーム" Then
                ws.Delete
            End If
        Next i
        
        '=====================================
        '  ★ 10_metadata_browser の処理
        '=====================================
        If Dir(browserPath) <> "" Then
            Set wbBrowser = Workbooks.Open(browserPath, ReadOnly:=True)
            Call SetBrowserDataToTable(wbBrowser, wsTable)
            wbBrowser.Close SaveChanges:=False
            Set wbBrowser = Nothing
        End If
        
        '▼ LogicalNameを取得（ファイル名生成用）
        '▼ SetBrowserDataToTable の後に確実にデータが書き込まれた状態で読み取る
        Application.Calculate
        DoEvents
        logicalName = GetLogicalNameFromBrowser(wsTable)
        If logicalName = "" Then
            Dim dotPos As Long
            dotPos = InStrRev(fileName, ".")
            If dotPos > 0 Then
                logicalName = Left(fileName, dotPos - 1)
            Else
                logicalName = fileName
            End If
        End If
        
        '▼ ファイル名から使用できない文字を削除
        logicalName = SanitizeFileName(logicalName)
        
        '=====================================
        '  ★ 30_attribute の処理
        '=====================================
        If Dir(attributePath) <> "" Then
            Set wbAttribute = Workbooks.Open(attributePath, ReadOnly:=True)
            Call SetAttributeDataToForm(wbAttribute, wsForm)
            wbAttribute.Close SaveChanges:=False
            Set wbAttribute = Nothing
        End If
        
        '=====================================
        '  ★ 20_metadata_document_generator の処理
        '=====================================
        If Dir(documentPath) <> "" Then
            Set wbDocument = Workbooks.Open(documentPath, ReadOnly:=False)
            Call SetDocumentDataToForm(wbDocument, wsForm)
            wbDocument.Close SaveChanges:=False
            Set wbDocument = Nothing
        End If
        
        '▼ 出力ファイル名を生成
        outputPath = folderOutput & logicalName & ".xlsx"
        
        '▼ 既存ファイルがある場合は上書き確認（エラー回避のため）
        If Dir(outputPath) <> "" Then
            Kill outputPath
        End If
        
        '▼ 保存して閉じる
        On Error Resume Next
        wbOut.SaveAs outputPath
        Dim saveErrNum As Long
        Dim saveErrDesc As String
        saveErrNum = Err.Number
        saveErrDesc = Err.Description
        Err.Clear
        On Error GoTo ERR_HANDLER
        
        If saveErrNum <> 0 Then
            Dim saveErrMsg As String
            saveErrMsg = "ファイル保存時にエラーが発生しました。" & vbCrLf & vbCrLf
            saveErrMsg = saveErrMsg & "ファイル名: " & logicalName & vbCrLf
            saveErrMsg = saveErrMsg & "パス: " & outputPath & vbCrLf
            saveErrMsg = saveErrMsg & "エラー番号: " & saveErrNum & vbCrLf
            saveErrMsg = saveErrMsg & "エラー内容: " & saveErrDesc
            MsgBox saveErrMsg, vbCritical, "保存エラー"
            wbOut.Close SaveChanges:=False
            Set wbOut = Nothing
            GoTo NEXT_FILE
        End If
        
        wbOut.Close SaveChanges:=False
        Set wbOut = Nothing
        
NEXT_FILE:
        fileName = Dir()
    Loop
    
    '▼ ループ終了後の処理
    On Error Resume Next
    Application.DisplayAlerts = True
    On Error GoTo ERR_HANDLER
    
    MsgBox "メタデータの結合が完了しました。", vbInformation
    
    Exit Sub

'────────────────────────────────────────
ERR_HANDLER:
'────────────────────────────────────────
    '▼ エラー情報を最初に保存（On Error Resume Next の前に）
    Dim errNum As Long
    Dim errDesc As String
    Dim errSource As String
    
    errNum = Err.Number
    errDesc = Err.Description
    errSource = Err.Source
    
    '▼ エラー情報を保存した後、エラーを無視してクリーンアップ処理
    On Error Resume Next
    Application.DisplayAlerts = False
    If Not wbOut Is Nothing Then wbOut.Close SaveChanges:=False
    If Not wbBrowser Is Nothing Then wbBrowser.Close SaveChanges:=False
    If Not wbDocument Is Nothing Then wbDocument.Close SaveChanges:=False
    If Not wbAttribute Is Nothing Then wbAttribute.Close SaveChanges:=False
    Application.DisplayAlerts = True
    On Error GoTo 0
    
    '▼ 保存したエラー情報を表示
    Dim errMsg As String
    errMsg = "エラーが発生しました。" & vbCrLf & vbCrLf
    errMsg = errMsg & "エラー番号：" & errNum & vbCrLf
    errMsg = errMsg & "エラー内容：" & errDesc
    If errSource <> "" Then
        errMsg = errMsg & vbCrLf & "エラー発生元：" & errSource
    End If
    
    MsgBox errMsg, vbCritical, "エラー"

End Sub


'========================================================================
'  10_metadata_browser のデータを「テーブル」シートに出力
'========================================================================
Private Sub SetBrowserDataToTable(wbBrowser As Workbook, wsTable As Worksheet)

    Dim wsBrowser As Worksheet
    Dim colIndex As Object
    Dim lastCol As Long
    Dim lastRow As Long
    Dim col As Long
    Dim row As Long
    Dim targetCols As Variant
    Dim i As Long
    Dim colName As String
    Dim colNum As Long
    Dim outputCol As Long
    
    Set wsBrowser = wbBrowser.Sheets(1)
    Set colIndex = CreateObject("Scripting.Dictionary")
    
    '▼ 抽出対象の列名リスト
    targetCols = Array("LogicalName", "SchemaName", "AutoCreateAccessTeams", _
                       "Change TrackingEnabled", "Description", "DisplayCollectionName", _
                       "DisplayName", "EntityColor", "EntityHelpUrl", "EntityHelpUrlEnabled", _
                       "HasActivities", "HasFeedback", "HasNotes", "IsAuditEnabled", _
                       "IsAvailableOffline", "IsConnectionsEnabled", "IsDocumentManagementEnabled", _
                       "IsDuplicateDetectionEnabled", "IsKnowledgeManagementEnabled", _
                       "IsMailMergeEnabled", "IsQuickCreateEnabled", "IsSLAEnabled", _
                       "IsValidF orAdvanced-ind", "IsValidForQueue", "OwnershipType", _
                       "Primarylmage Attribute", "TableType")
    
    '▼ 列名のインデックスを取得
    lastCol = wsBrowser.Cells(1, wsBrowser.Columns.Count).End(xlToLeft).Column
    
    For col = 1 To lastCol
        colName = Trim(wsBrowser.Cells(1, col).Value)
        If colName <> "" Then
            colIndex(colName) = col
        End If
    Next col
    
    '▼ データ行数を取得
    lastRow = wsBrowser.Cells(wsBrowser.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then Exit Sub  'データ行がない場合は終了
    
    '▼ ヘッダー行を出力（A1から）
    outputCol = 1
    For i = LBound(targetCols) To UBound(targetCols)
        colName = targetCols(i)
        wsTable.Cells(1, outputCol).Value = colName
        outputCol = outputCol + 1
    Next i
    
    '▼ データ行を出力（2行目から）
    For row = 2 To lastRow
        outputCol = 1
        For i = LBound(targetCols) To UBound(targetCols)
            colName = targetCols(i)
            Dim cellValue As String
            If colIndex.Exists(colName) Then
                colNum = colIndex(colName)
                cellValue = Trim(wsBrowser.Cells(row, colNum).Value)
            Else
                cellValue = ""
            End If
            
            '▼ 値の変換処理を適用
            cellValue = ConvertTableValue(colName, cellValue)
            wsTable.Cells(row, outputCol).Value = cellValue
            outputCol = outputCol + 1
        Next i
    Next row
    
    '▼ データの書き込みを確実に完了させる
    Application.Calculate
    DoEvents

End Sub


'========================================================================
'  30_attribute のデータを「フォーム」シートに出力
'========================================================================
Private Sub SetAttributeDataToForm(wbAttribute As Workbook, wsForm As Worksheet)

    Dim wsAttribute As Worksheet
    Dim colIndex As Object
    Dim lastCol As Long
    Dim lastRow As Long
    Dim col As Long
    Dim row As Long
    Dim targetCols As Variant
    Dim i As Long
    Dim colName As String
    Dim colNum As Long
    Dim outputCol As Long
    
    Set wsAttribute = wbAttribute.Sheets(1)
    Set colIndex = CreateObject("Scripting.Dictionary")
    
    '▼ 抽出対象の列名リスト
    targetCols = Array("SchemaName", "DisplayName", "LogicalName", "CustomAttribute", _
                       "IsCustomAttribute", "IsPrimaryID", "IsPrimaryName", _
                       "AttributeTypeName", "RequiredLevel", "IsAuditEnabled", _
                       "IsGlobalFilterEnabled", "IsSortableEnabled", "IsSearchable", _
                       "Description", "IsSecret")
    
    '▼ 列名のインデックスを取得
    lastCol = wsAttribute.Cells(1, wsAttribute.Columns.Count).End(xlToLeft).Column
    
    For col = 1 To lastCol
        colName = Trim(wsAttribute.Cells(1, col).Value)
        If colName <> "" Then
            colIndex(colName) = col
        End If
    Next col
    
    '▼ データ行数を取得
    lastRow = wsAttribute.Cells(wsAttribute.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then Exit Sub  'データ行がない場合は終了
    
    '▼ ヘッダー行を出力（A1から）
    outputCol = 1
    For i = LBound(targetCols) To UBound(targetCols)
        colName = targetCols(i)
        wsForm.Cells(1, outputCol).Value = colName
        outputCol = outputCol + 1
    Next i
    
    '▼ 追加列のヘッダーを出力（既存の列の後ろに）
    Dim additionalColsHeader As Variant
    additionalColsHeader = Array("最大値", "最小値", "小数点以下表示桁数", "通貨の精度", _
                                 "タイムゾーン", "選択肢", "規定値", "関連テーブル", "リレーションシップ名列", "文字数列")
    Dim j As Long
    For j = LBound(additionalColsHeader) To UBound(additionalColsHeader)
        wsForm.Cells(1, outputCol).Value = additionalColsHeader(j)
        outputCol = outputCol + 1
    Next j
    
    '▼ データ行を出力（2行目から）
    Dim outputRow As Long
    outputRow = 1  'ヘッダー行の次の行から開始
    
    For row = 2 To lastRow
        '▼ DisplayNameを取得してN/Aチェック
        Dim displayNameValue As String
        If colIndex.Exists("DisplayName") Then
            displayNameValue = Trim(wsAttribute.Cells(row, colIndex("DisplayName")).Value)
        Else
            displayNameValue = ""
        End If
        
        '▼ DisplayNameがN/Aの行はスキップ
        If LCase(displayNameValue) = "n/a" Then
            GoTo NEXT_ATTRIBUTE_ROW
        End If
        
        outputRow = outputRow + 1
        outputCol = 1
        For i = LBound(targetCols) To UBound(targetCols)
            colName = targetCols(i)
            Dim cellValue As String
            If colIndex.Exists(colName) Then
                colNum = colIndex(colName)
                cellValue = Trim(wsAttribute.Cells(row, colNum).Value)
            Else
                cellValue = ""
            End If
            
            '▼ 値の変換処理を適用
            cellValue = ConvertFormValue(colName, cellValue)
            wsForm.Cells(outputRow, outputCol).Value = cellValue
            outputCol = outputCol + 1
        Next i
        
        '▼ 追加列を出力（空のセルとして）
        Dim additionalCols As Variant
        additionalCols = Array("最大値", "最小値", "小数点以下表示桁数", "通貨の精度", _
                                 "タイムゾーン", "選択肢", "規定値", "関連テーブル", "リレーションシップ名列", "文字数列")
        For j = LBound(additionalCols) To UBound(additionalCols)
            wsForm.Cells(outputRow, outputCol).Value = ""
            outputCol = outputCol + 1
        Next j
        
NEXT_ATTRIBUTE_ROW:
    Next row

End Sub


'========================================================================
'  20_metadata_document_generator のデータを「フォーム」シートに追加
'========================================================================
Private Sub SetDocumentDataToForm(wbDocument As Workbook, wsForm As Worksheet)

    Dim wsDocument As Worksheet
    Dim colIndex As Object
    Dim formColIndex As Object
    Dim lastCol As Long
    Dim lastRow As Long
    Dim formLastRow As Long
    Dim col As Long
    Dim row As Long
    Dim formRow As Long
    Dim targetCols As Variant
    Dim i As Long
    Dim colName As String
    Dim colNum As Long
    Dim logicalNameCol As Long
    Dim formLogicalNameCol As Long
    Dim logicalNameValue As String
    Dim formLogicalNameValue As String
    Dim found As Boolean
    Dim maxIterations As Long
    Dim iterationCount As Long
    
    Set wsDocument = wbDocument.Sheets(1)
    Set colIndex = CreateObject("Scripting.Dictionary")
    Set formColIndex = CreateObject("Scripting.Dictionary")
    
    '▼ 抽出対象の列名リスト
    targetCols = Array("Form location", "Additional data", "Type")
    
    '▼ 列名のインデックスを取得（document）
    lastCol = wsDocument.Cells(1, wsDocument.Columns.Count).End(xlToLeft).Column
    
    For col = 1 To lastCol
        colName = Trim(wsDocument.Cells(1, col).Value)
        If colName <> "" Then
            colIndex(colName) = col
        End If
    Next col
    
    '▼ 「Logical Name」列の位置を取得（document）
    If Not colIndex.Exists("Logical Name") Then
        Exit Sub  '「Logical Name」列がない場合は終了
    End If
    logicalNameCol = colIndex("Logical Name")
    
    '▼ フォームシートの列名のインデックスを取得
    lastCol = wsForm.Cells(1, wsForm.Columns.Count).End(xlToLeft).Column
    For col = 1 To lastCol
        colName = Trim(wsForm.Cells(1, col).Value)
        If colName <> "" Then
            formColIndex(colName) = col
        End If
    Next col
    
    '▼ 「LogicalName」列の位置を取得（form）
    If Not formColIndex.Exists("LogicalName") Then
        Exit Sub  '「LogicalName」列がない場合は終了
    End If
    formLogicalNameCol = formColIndex("LogicalName")
    
    '▼ フォームシートの最終行を取得
    formLastRow = wsForm.Cells(wsForm.Rows.Count, 1).End(xlUp).Row
    
    '▼ 追加する列のヘッダーを追加（既存の列の後ろに）
    Dim nextCol As Long
    nextCol = lastCol + 1
    For i = LBound(targetCols) To UBound(targetCols)
        colName = targetCols(i)
        wsForm.Cells(1, nextCol).Value = colName
        nextCol = nextCol + 1
    Next i
    
    '▼ データ行数を取得（document）
    lastRow = wsDocument.Cells(wsDocument.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then Exit Sub  'データ行がない場合は終了
    
    '▼ 無限ループ防止：最大処理回数を設定
    maxIterations = 10000
    iterationCount = 0
    
    '▼ documentの各行について、フォームシートで一致する行を探して追加
    For row = 2 To lastRow
        iterationCount = iterationCount + 1
        If iterationCount > maxIterations Then Exit For
        
        logicalNameValue = Trim(wsDocument.Cells(row, logicalNameCol).Value)
        If logicalNameValue = "" Then GoTo NEXT_ROW
        
        found = False
        
        '▼ フォームシートで一致する行を探す
        For formRow = 2 To formLastRow
            formLogicalNameValue = Trim(wsForm.Cells(formRow, formLogicalNameCol).Value)
            If LCase(logicalNameValue) = LCase(formLogicalNameValue) Then
                found = True
                
                '▼ 一致した行にデータを追加
                nextCol = lastCol + 1
                Dim additionalDataValue As String
                additionalDataValue = ""
                
                For i = LBound(targetCols) To UBound(targetCols)
                    colName = targetCols(i)
                    Dim docCellValue As String
                    If colIndex.Exists(colName) Then
                        colNum = colIndex(colName)
                        docCellValue = Trim(wsDocument.Cells(row, colNum).Value)
                    Else
                        docCellValue = ""
                    End If
                    
                    '▼ Additional dataを保存
                    If colName = "Additional data" Then
                        additionalDataValue = docCellValue
                    End If
                    
                    '▼ Typeの変換処理を適用
                    If colName = "Type" Then
                        docCellValue = ConvertFormType(docCellValue)
                    End If
                    
                    wsForm.Cells(formRow, nextCol).Value = docCellValue
                    nextCol = nextCol + 1
                Next i
                
                '▼ Additional dataのテキストを解析して対応する列に値をセット
                If additionalDataValue <> "" Then
                    Call ParseAdditionalData(wsForm, formRow, additionalDataValue, formColIndex)
                End If
                
                Exit For
            End If
        Next formRow
        
NEXT_ROW:
    Next row

End Sub


'========================================================================
'  browser から LogicalName を取得
'========================================================================
Private Function GetLogicalNameFromBrowser(wsTable As Worksheet) As String

    Dim lastCol As Long
    Dim col As Long
    Dim colName As String
    Dim retryCount As Long
    Dim maxRetries As Long
    
    maxRetries = 3
    
    '▼ データが正しく書き込まれるまで最大3回再試行
    For retryCount = 1 To maxRetries
        '▼ Excelの処理を完了させる
        Application.Calculate
        DoEvents
        
        '▼ 列数を取得
        lastCol = wsTable.Cells(1, wsTable.Columns.Count).End(xlToLeft).Column
        
        '▼ データが書き込まれているか確認（列数が1以上、かつA1セルに値がある）
        If lastCol > 0 And Trim(wsTable.Cells(1, 1).Value) <> "" Then
            '▼ LogicalName列を探す
            For col = 1 To lastCol
                colName = Trim(wsTable.Cells(1, col).Value)
                If LCase(colName) = "logicalname" Then
                    Dim logicalNameValue As String
                    logicalNameValue = Trim(wsTable.Cells(2, col).Value)
                    If logicalNameValue <> "" Then
                        GetLogicalNameFromBrowser = logicalNameValue
                        Exit Function
                    End If
                End If
            Next col
        End If
        
        '▼ データがまだ書き込まれていない場合は少し待つ
        If retryCount < maxRetries Then
            Dim waitCount As Long
            For waitCount = 1 To 10
                DoEvents
            Next waitCount
        End If
    Next retryCount
    
    GetLogicalNameFromBrowser = ""

End Function


'========================================================================
'  ファイル名から使用できない文字を削除
'========================================================================
Private Function SanitizeFileName(fileName As String) As String

    Dim result As String
    Dim i As Long
    Dim char As String
    Dim invalidChars As String
    
    '▼ Windowsで使用できない文字
    invalidChars = "/\:*?""<>|"
    
    result = fileName
    
    '▼ 使用できない文字を削除
    For i = 1 To Len(invalidChars)
        char = Mid(invalidChars, i, 1)
        result = Replace(result, char, "_")
    Next i
    
    '▼ 先頭・末尾のスペースとピリオドを削除
    result = Trim(result)
    While Right(result, 1) = "." Or Right(result, 1) = " "
        result = Left(result, Len(result) - 1)
    Wend
    
    '▼ 空文字の場合はデフォルト名を返す
    If result = "" Then
        result = "output"
    End If
    
    '▼ ファイル名の長さ制限（255文字）
    If Len(result) > 255 Then
        result = Left(result, 255)
    End If
    
    SanitizeFileName = result

End Function


'========================================================================
'  テーブル値の変換（True/False・所有権・種類・画像）
'========================================================================
Private Function ConvertTableValue(key As String, val As String) As String
    
    val = Trim(val)
    
    Select Case key

        Case "TableType"
            Select Case val
                Case "Standard": ConvertTableValue = "標準"
                Case "Activity": ConvertTableValue = "活動"
                Case "Virtual": ConvertTableValue = "仮想"
                Case Else: ConvertTableValue = val
            End Select
        
        Case "OwnershipType"
            Select Case val
                Case "UserOwned": ConvertTableValue = "ユーザーまたはチーム"
                Case "OrganizationOwned": ConvertTableValue = "組織"
                Case Else: ConvertTableValue = val
            End Select
        
        Case "Primarylmage Attribute", "PrimaryImageAttribute"
            If val = "" Then
                ConvertTableValue = "なし"
            Else
                ConvertTableValue = "あり"
            End If
        
        Case "EntityColor"
            If val = "0" Or val = 0 Then
                ConvertTableValue = "-"
            Else
                ConvertTableValue = val
            End If
        
        Case Else
            '▼ True/Falseの値をチェック/- に変換
            If LCase(val) = "true" Then
                ConvertTableValue = ChrW(10003)  'チェックマーク（✓）
            ElseIf LCase(val) = "false" Or val = "" Then
                ConvertTableValue = "-"
            Else
                ConvertTableValue = val
            End If
    End Select

End Function


'========================================================================
'  フォーム値の変換（IsPrimaryName・RequiredLevel）
'========================================================================
Private Function ConvertFormValue(key As String, val As String) As String
    
    val = Trim(val)
    
    Select Case key
        
        Case "IsPrimaryName"
            '▼ True → ○ (e2 97 8b, ChrW(9675))、False → -
            If LCase(val) = "true" Then
                ConvertFormValue = ChrW(9675)  '○
            ElseIf LCase(val) = "false" Or val = "" Then
                ConvertFormValue = "-"
            Else
                ConvertFormValue = val
            End If
        
        Case "RequiredLevel"
            '▼ None → -、SystemRequired → 必須項目
            Select Case val
                Case "None": ConvertFormValue = "任意"
                Case "SystemRequired": ConvertFormValue = "必須項目"
                Case Else: ConvertFormValue = val
            End Select
        
        Case Else
            '▼ その他のTrue/Falseはそのまま
            ConvertFormValue = val
    End Select

End Function


'========================================================================
'  フォームTypeの変換
'========================================================================
Private Function ConvertFormType(val As String) As String
    
    val = Trim(val)
    
    Select Case val
        Case "Simple": ConvertFormType = "シンプル"
        Case "Calculated": ConvertFormType = "計算済みの列"
        Case "Rollup": ConvertFormType = "ロールアップ列"
        Case Else: ConvertFormType = val
    End Select

End Function


'========================================================================
'  Additional dataのテキストを解析して対応する列に値をセット
'========================================================================
Private Sub ParseAdditionalData(wsForm As Worksheet, formRow As Long, additionalData As String, formColIndex As Object)
    
    Dim keywordPos As Long
    Dim valueText As String
    Dim colNum As Long
    Dim nextKeywordPos As Long
    Dim valueStart As Long
    Dim valueEnd As Long
    
    '▼ 全体のテキストから各キーワードを探して値を取得
    
    '▼ Targets:が含まれている場合
    keywordPos = InStr(1, additionalData, "Targets:", vbTextCompare)
    If keywordPos > 0 Then
        valueStart = keywordPos + Len("Targets:")
        '▼ 次のキーワードの位置を探す
        nextKeywordPos = FindNextKeyword(additionalData, valueStart)
        If nextKeywordPos > 0 Then
            valueText = Trim(Mid(additionalData, valueStart, nextKeywordPos - valueStart))
        Else
            valueText = Trim(Mid(additionalData, valueStart))
        End If
        '▼ 改行をスペースに置換
        valueText = Replace(valueText, vbCrLf, " ")
        valueText = Replace(valueText, vbLf, " ")
        valueText = Replace(valueText, vbCr, " ")
        valueText = Trim(valueText)
        If formColIndex.Exists("関連テーブル") Then
            colNum = formColIndex("関連テーブル")
            wsForm.Cells(formRow, colNum).Value = valueText
        End If
    End If
    
    '▼ Format:が含まれている場合（Format: Textは除外）
    keywordPos = InStr(1, additionalData, "Format:", vbTextCompare)
    If keywordPos > 0 Then
        '▼ Format: Textの場合は除外
        Dim checkText As String
        checkText = Mid(additionalData, keywordPos, Len("Format: Text"))
        If LCase(checkText) <> LCase("Format: Text") Then
            '▼ Format: Textではない場合のみ処理
            valueStart = keywordPos + Len("Format:")
            nextKeywordPos = FindNextKeyword(additionalData, valueStart)
            If nextKeywordPos > 0 Then
                valueText = Trim(Mid(additionalData, valueStart, nextKeywordPos - valueStart))
            Else
                valueText = Trim(Mid(additionalData, valueStart))
            End If
            '▼ 改行をスペースに置換
            valueText = Replace(valueText, vbCrLf, " ")
            valueText = Replace(valueText, vbLf, " ")
            valueText = Replace(valueText, vbCr, " ")
            valueText = Trim(valueText)
            If formColIndex.Exists("タイムゾーン") Then
                colNum = formColIndex("タイムゾーン")
                wsForm.Cells(formRow, colNum).Value = valueText
            End If
        End If
    End If
    
    '▼ States:が含まれている場合
    keywordPos = InStr(1, additionalData, "States:", vbTextCompare)
    If keywordPos > 0 Then
        valueStart = keywordPos + Len("States:")
        nextKeywordPos = FindNextKeyword(additionalData, valueStart)
        If nextKeywordPos > 0 Then
            valueText = Trim(Mid(additionalData, valueStart, nextKeywordPos - valueStart))
        Else
            valueText = Trim(Mid(additionalData, valueStart))
        End If
        '▼ 改行をスペースに置換
        valueText = Replace(valueText, vbCrLf, " ")
        valueText = Replace(valueText, vbLf, " ")
        valueText = Replace(valueText, vbCr, " ")
        valueText = Trim(valueText)
        If formColIndex.Exists("選択肢") Then
            colNum = formColIndex("選択肢")
            wsForm.Cells(formRow, colNum).Value = valueText
        End If
    End If
    
    '▼ Format: Textが含まれている場合、Max length:以降の数値を取得
    keywordPos = InStr(1, additionalData, "Format: Text", vbTextCompare)
    If keywordPos > 0 Then
        Dim maxLengthPos As Long
        maxLengthPos = InStr(keywordPos, additionalData, "Max length:", vbTextCompare)
        If maxLengthPos > 0 Then
            valueStart = maxLengthPos + Len("Max length:")
            nextKeywordPos = FindNextKeyword(additionalData, valueStart)
            Dim maxLengthText As String
            If nextKeywordPos > 0 Then
                maxLengthText = Trim(Mid(additionalData, valueStart, nextKeywordPos - valueStart))
            Else
                maxLengthText = Trim(Mid(additionalData, valueStart))
            End If
            '▼ 数値部分を抽出
            Dim numValue As String
            numValue = ExtractNumber(maxLengthText)
            '▼ 文字数列の列にセット
            If formColIndex.Exists("文字数列") Then
                colNum = formColIndex("文字数列")
                wsForm.Cells(formRow, colNum).Value = numValue
            End If
        End If
    End If
    
    '▼ Minimum value:が含まれている場合
    keywordPos = InStr(1, additionalData, "Minimum value:", vbTextCompare)
    If keywordPos > 0 Then
        valueStart = keywordPos + Len("Minimum value:")
        nextKeywordPos = FindNextKeyword(additionalData, valueStart)
        If nextKeywordPos > 0 Then
            valueText = Trim(Mid(additionalData, valueStart, nextKeywordPos - valueStart))
        Else
            valueText = Trim(Mid(additionalData, valueStart))
        End If
        '▼ 改行をスペースに置換
        valueText = Replace(valueText, vbCrLf, " ")
        valueText = Replace(valueText, vbLf, " ")
        valueText = Replace(valueText, vbCr, " ")
        valueText = Trim(valueText)
        If formColIndex.Exists("最小値") Then
            colNum = formColIndex("最小値")
            wsForm.Cells(formRow, colNum).Value = valueText
        End If
    End If
    
    '▼ Maximum value:が含まれている場合
    keywordPos = InStr(1, additionalData, "Maximum value:", vbTextCompare)
    If keywordPos > 0 Then
        valueStart = keywordPos + Len("Maximum value:")
        nextKeywordPos = FindNextKeyword(additionalData, valueStart)
        If nextKeywordPos > 0 Then
            valueText = Trim(Mid(additionalData, valueStart, nextKeywordPos - valueStart))
        Else
            valueText = Trim(Mid(additionalData, valueStart))
        End If
        '▼ 改行をスペースに置換
        valueText = Replace(valueText, vbCrLf, " ")
        valueText = Replace(valueText, vbLf, " ")
        valueText = Replace(valueText, vbCr, " ")
        valueText = Trim(valueText)
        If formColIndex.Exists("最大値") Then
            colNum = formColIndex("最大値")
            wsForm.Cells(formRow, colNum).Value = valueText
        End If
    End If

End Sub


'========================================================================
'  次のキーワードの位置を探す
'========================================================================
Private Function FindNextKeyword(text As String, startPos As Long) As Long
    
    Dim keywords As Variant
    Dim keyword As Variant
    Dim pos As Long
    Dim minPos As Long
    
    keywords = Array("Targets:", "Format:", "States:", "Max length:", "Minimum value:", "Maximum value:")
    minPos = 0
    
    For Each keyword In keywords
        pos = InStr(startPos, text, keyword, vbTextCompare)
        If pos > 0 Then
            If minPos = 0 Or pos < minPos Then
                minPos = pos
            End If
        End If
    Next keyword
    
    FindNextKeyword = minPos

End Function


'========================================================================
'  テキストから数値を抽出
'========================================================================
Private Function ExtractNumber(text As String) As String
    
    Dim result As String
    Dim i As Long
    Dim char As String
    Dim foundDigit As Boolean
    
    result = ""
    foundDigit = False
    
    For i = 1 To Len(text)
        char = Mid(text, i, 1)
        If IsNumeric(char) Then
            result = result & char
            foundDigit = True
        ElseIf foundDigit Then
            '▼ 数値の後に非数値文字が来たら終了
            Exit For
        End If
    Next i
    
    ExtractNumber = result

End Function

