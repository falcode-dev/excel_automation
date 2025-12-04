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
    
    Dim displayName As String
    
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
    
    '▼ 無限ループ防止：最大処理回数を設定
    Dim maxIterations As Long
    Dim iterationCount As Long
    maxIterations = 1000
    iterationCount = 0
    
    Do While fileName <> "" And iterationCount < maxIterations
        iterationCount = iterationCount + 1
        
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
        
        '▼ デフォルトシートの名前を変更
        Dim wsTable As Worksheet
        Dim wsForm As Worksheet
        
        '▼ 最初のシートを「テーブル」に変更
        If wbOut.Sheets.Count >= 1 Then
            Set wsTable = wbOut.Sheets(1)
            wsTable.Name = "テーブル"
        Else
            Set wsTable = wbOut.Sheets.Add
            wsTable.Name = "テーブル"
        End If
        
        '▼ 2番目のシートを「フォーム」に変更（存在しない場合は作成）
        If wbOut.Sheets.Count >= 2 Then
            Set wsForm = wbOut.Sheets(2)
            wsForm.Name = "フォーム"
        Else
            Set wsForm = wbOut.Sheets.Add
            wsForm.Name = "フォーム"
        End If
        
        '▼ 3つ目以降のシートがあれば削除
        While wbOut.Sheets.Count > 2
            wbOut.Sheets(3).Delete
        Wend
        
        '=====================================
        '  ★ 10_metadata_browser の処理
        '=====================================
        If Dir(browserPath) <> "" Then
            Set wbBrowser = Workbooks.Open(browserPath, ReadOnly:=True)
            Call SetBrowserDataToTable(wbBrowser, wsTable)
            wbBrowser.Close SaveChanges:=False
            Set wbBrowser = Nothing
        End If
        
        '▼ DisplayNameを取得（ファイル名生成用）
        displayName = GetDisplayNameFromBrowser(wsTable)
        If displayName = "" Then
            Dim dotPos As Long
            dotPos = InStrRev(fileName, ".")
            If dotPos > 0 Then
                displayName = Left(fileName, dotPos - 1)
            Else
                displayName = fileName
            End If
        End If
        
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
        outputPath = folderOutput & displayName & ".xlsx"
        
        '▼ 既存ファイルがある場合は上書き確認（エラー回避のため）
        If Dir(outputPath) <> "" Then
            Kill outputPath
        End If
        
        '▼ 保存して閉じる
        wbOut.SaveAs outputPath
        wbOut.Close SaveChanges:=False
        Set wbOut = Nothing
        
        Application.DisplayAlerts = True
        
NEXT_FILE:
        fileName = Dir()
    Loop
    
    If iterationCount >= maxIterations Then
        MsgBox "警告: 最大処理回数(" & maxIterations & ")に達しました。処理を中断しました。", vbExclamation
    Else
        MsgBox "メタデータの結合が完了しました。", vbInformation
    End If
    
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
            If colIndex.Exists(colName) Then
                colNum = colIndex(colName)
                wsTable.Cells(row, outputCol).Value = Trim(wsBrowser.Cells(row, colNum).Value)
            Else
                wsTable.Cells(row, outputCol).Value = ""
            End If
            outputCol = outputCol + 1
        Next i
    Next row

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
    
    '▼ データ行を出力（2行目から）
    For row = 2 To lastRow
        outputCol = 1
        For i = LBound(targetCols) To UBound(targetCols)
            colName = targetCols(i)
            If colIndex.Exists(colName) Then
                colNum = colIndex(colName)
                wsForm.Cells(row, outputCol).Value = Trim(wsAttribute.Cells(row, colNum).Value)
            Else
                wsForm.Cells(row, outputCol).Value = ""
            End If
            outputCol = outputCol + 1
        Next i
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
                For i = LBound(targetCols) To UBound(targetCols)
                    colName = targetCols(i)
                    If colIndex.Exists(colName) Then
                        colNum = colIndex(colName)
                        wsForm.Cells(formRow, nextCol).Value = Trim(wsDocument.Cells(row, colNum).Value)
                    End If
                    nextCol = nextCol + 1
                Next i
                
                Exit For
            End If
        Next formRow
        
NEXT_ROW:
    Next row

End Sub


'========================================================================
'  browser から DisplayName を取得
'========================================================================
Private Function GetDisplayNameFromBrowser(wsTable As Worksheet) As String

    Dim lastCol As Long
    Dim col As Long
    Dim colName As String
    
    lastCol = wsTable.Cells(1, wsTable.Columns.Count).End(xlToLeft).Column
    
    For col = 1 To lastCol
        colName = Trim(wsTable.Cells(1, col).Value)
        If LCase(colName) = "displayname" Then
            GetDisplayNameFromBrowser = Trim(wsTable.Cells(2, col).Value)
            Exit Function
        End If
    Next col
    
    GetDisplayNameFromBrowser = ""

End Function

