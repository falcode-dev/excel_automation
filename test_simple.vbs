Option Explicit

'────────────────────────────────────────
'  メイン処理（シンプル版：テンプレート不使用）
'────────────────────────────────────────
Public Sub メイン処理_Entityのみ()

    Dim wbThis As Workbook
    Dim folderBase As String
    Dim folderEntity As String
    Dim folderAttribute As String
    Dim folderOutput As String
    Dim folderTemplate As String
    Dim templatePath As String
    
    Dim entityFile As String
    Dim entityPath As String
    Dim attributePath As String
    Dim outputPath As String
    
    Dim wbOut As Workbook
    Dim wbEntity As Workbook
    Dim wbAttribute As Workbook
    Dim wbTemplate As Workbook
    
    On Error GoTo ERR_HANDLER
    
    Set wbThis = ThisWorkbook
    folderBase = wbThis.Path & "\"
    
    '▼ 必要なフォルダ
    folderEntity = folderBase & "10_entity\"
    folderAttribute = folderBase & "20_attribute\"
    folderOutput = folderBase & "30_create_file\"
    folderTemplate = folderBase & "template\"
    
    '▼ フォルダ存在チェック
    If Dir(folderEntity, vbDirectory) = "" Then Err.Raise 100, , "10_entity フォルダがありません。"
    If Dir(folderAttribute, vbDirectory) = "" Then Err.Raise 105, , "20_attribute フォルダがありません。"
    If Dir(folderOutput, vbDirectory) = "" Then Err.Raise 102, , "30_create_file フォルダがありません。"
    If Dir(folderTemplate, vbDirectory) = "" Then Err.Raise 101, , "template フォルダがありません。"
    templatePath = folderTemplate & "template.xlsx"
    If Dir(templatePath) = "" Then Err.Raise 103, , "template.xlsx が見つかりません。"
    
    '▼ entity フォルダの全Excelを処理
    entityFile = Dir(folderEntity & "*.xlsx")
    If entityFile = "" Then Err.Raise 104, , "10_entity に処理対象ファイルがありません。"
    
    Do While entityFile <> ""
    
        entityPath = folderEntity & entityFile
        
        '▼ entity を開く
        Set wbEntity = Workbooks.Open(entityPath, ReadOnly:=True)
        
        '▼ 表示名を取得してファイル名を生成（エンティティ定義書_ID_XXX_v0.0.xlsx）
        Dim displayName As String
        displayName = GetDisplayNameFromEntity(wbEntity)
        If displayName = "" Then
            '表示名が取得できない場合は元のファイル名を使用
            Dim dotPos As Long
            dotPos = InStrRev(entityFile, ".")
            If dotPos > 0 Then
                displayName = Left(entityFile, dotPos - 1)
            Else
                displayName = entityFile
            End If
        End If
        outputPath = folderOutput & "エンティティ定義書_ID_" & displayName & "_v0.0.xlsx"
        
        '▼ テンプレートを開く
        Set wbTemplate = Workbooks.Open(templatePath, ReadOnly:=True)
        
        '▼ テンプレートをコピー
        wbTemplate.SaveCopyAs outputPath
        Set wbOut = Workbooks.Open(outputPath)
        
        '=====================================
        '  ★ entity の情報をテンプレートに出力
        '=====================================
        Call SetEntityInfoToTemplate(wbOut, wbEntity)
        
        '▼ entity ファイルを閉じる
        wbEntity.Close SaveChanges:=False
        
        '▼ attribute ファイルを開く（存在する場合）
        attributePath = folderAttribute & entityFile
        If Dir(attributePath) <> "" Then
            Set wbAttribute = Workbooks.Open(attributePath, ReadOnly:=True)
            
            '=====================================
            '  ★ attribute の情報をテンプレートに出力
            '=====================================
            Call SetAttributeInfoToTemplate(wbOut, wbAttribute)
            
            wbAttribute.Close SaveChanges:=False
        End If
        
        '▼ 保存して閉じる
        wbOut.Close SaveChanges:=True
        wbTemplate.Close SaveChanges:=False
        
        entityFile = Dir()
    Loop
    
    MsgBox "entity データの出力が完了しました。", vbInformation
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
    If Not wbOut Is Nothing Then wbOut.Close SaveChanges:=False
    If Not wbTemplate Is Nothing Then wbTemplate.Close SaveChanges:=False
    If Not wbEntity Is Nothing Then wbEntity.Close SaveChanges:=False
    If Not wbAttribute Is Nothing Then wbAttribute.Close SaveChanges:=False
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
'  英語名 → 日本語名 辞書
'========================================================================
Private Function GetEntityMappingDict() As Object

    Dim dic As Object: Set dic = CreateObject("Scripting.Dictionary")

    dic("LogicalName") = "論理名"
    dic("SchemaName") = "スキーマ名"
    dic("AutoCreateAccessTeams") = "アクセスチームを有する"
    dic("ChangeTrackingEnabled") = "変更を追跡"
    dic("Description") = "説明"
    dic("DisplayCollectionName") = "複数形の名前"
    dic("DisplayName") = "表示名"
    dic("EntityColor") = "色"
    dic("EntityHelpUrl") = "ヘルプのURL"
    dic("EntityHelpUrlEnabled") = "カスタムヘルプを提供する"
    dic("HasActivities") = "活動の新規作成時"
    dic("HasFeedback") = "フィードバックにリンク可能"
    dic("HasNotes") = "添付ファイルを有効にする（メモファイルを含む）"
    dic("IsAuditEnabled") = "データに対する変更を監査する"
    dic("IsAvailableOffline") = "オフラインで取得できる"
    dic("IsConnectionsEnabled") = "接続可能"
    dic("IsDocumentManagementEnabled") = "SharePointドキュメント管理の設定時"
    dic("IsDuplicateDetectionEnabled") = "重複データ検出ルールの適用"
    dic("IsKnowledgeManagementEnabled") = "ナレッジマネージメントを許可する"
    dic("IsMailMergeEnabled") = "差し込み印刷時"
    dic("IsQuickCreateEnabled") = "可能な場合は簡易作成フォームを活用します"
    dic("IsSLAEnabled") = "サービスレベルアグリーメント設定時"
    dic("IsValidForAdvancedFind") = "検索結果に表示"
    dic("IsValidForQueue") = "キューに追加可能"
    dic("OwnershipType") = "所有権を記録する"
    dic("PrimaryImageAttribute") = "テーブルの画像を選択する"
    dic("TableType") = "種類"

    Set GetEntityMappingDict = dic
End Function


'========================================================================
'  entity から表示名を取得
'========================================================================
Private Function GetDisplayNameFromEntity(wbEntity As Workbook) As String

    Dim wsEntity As Worksheet
    Dim lastCol As Long
    Dim col As Long
    Dim engKey As String
    Dim val As String
    
    Set wsEntity = wbEntity.Sheets(1)
    lastCol = wsEntity.Cells(1, wsEntity.Columns.Count).End(xlToLeft).Column
    
    For col = 1 To lastCol
        engKey = Trim(wsEntity.Cells(1, col).Value)
        If LCase(engKey) = "displayname" Then
            val = Trim(wsEntity.Cells(2, col).Value)
            GetDisplayNameFromEntity = val
            Exit Function
        End If
    Next col
    
    GetDisplayNameFromEntity = ""

End Function


'========================================================================
'  値変換（True/False・所有権・種類・画像）
'========================================================================
Private Function ConvertEntityValue(key As String, val As String) As String
    
    val = Trim(val)
    
    Select Case key

        Case "TableType"
            Select Case val
                Case "Standard": ConvertEntityValue = "標準"
                Case "Activity": ConvertEntityValue = "活動"
                Case "Virtual": ConvertEntityValue = "仮想"
                Case Else: ConvertEntityValue = val
            End Select
        
        Case "OwnershipType"
            Select Case val
                Case "UserOwned": ConvertEntityValue = "ユーザーまたはチーム"
                Case "OrganizationOwned": ConvertEntityValue = "組織"
                Case Else: ConvertEntityValue = val
            End Select
        
        Case "PrimaryImageAttribute"
            If val = "" Then
                ConvertEntityValue = "なし"
            Else
                ConvertEntityValue = "あり"
            End If
        
        Case Else
            If LCase(val) = "true" Then
                ConvertEntityValue = "ON"
            ElseIf LCase(val) = "false" Then
                ConvertEntityValue = "-"
            Else
                ConvertEntityValue = val
            End If
    End Select

End Function


'========================================================================
'  entity.xlsx → テンプレートへ出力
'
'  出力先：テンプレートの指定セル
'  変更したセルは赤文字にする
'========================================================================
Public Sub SetEntityInfoToTemplate(wbOut As Workbook, wbEntity As Workbook)

    Dim wsEntity As Worksheet
    Dim wsCover As Worksheet
    Dim wsTable As Worksheet
    Dim dic As Object
    Dim lastCol As Long
    Dim col As Long
    Dim engKey As String, val As String
    Dim cell As Range
    
    Set wsEntity = wbEntity.Sheets(1)
    Set wsCover = wbOut.Sheets("表紙")
    Set wsTable = wbOut.Sheets("テーブル")
    Set dic = GetEntityMappingDict()
    
    '▼ entity からデータを取得
    Dim entityData As Object
    Set entityData = CreateObject("Scripting.Dictionary")
    
    lastCol = wsEntity.Cells(1, wsEntity.Columns.Count).End(xlToLeft).Column
    
    For col = 1 To lastCol
        engKey = Trim(wsEntity.Cells(1, col).Value)
        val = Trim(wsEntity.Cells(2, col).Value)
        If engKey <> "" Then
            entityData(engKey) = val
        End If
    Next col
    
    '▼ シート「表紙」に値を設定
    If entityData.Exists("DisplayName") Then
        Set cell = wsCover.Range("W21")
        cell.Value = ConvertEntityValue("DisplayName", entityData("DisplayName"))
        cell.Font.Color = RGB(255, 0, 0)  '赤文字
    End If
    
    '▼ シート「テーブル」に値を設定
    If entityData.Exists("DisplayName") Then
        Set cell = wsTable.Range("E5")
        cell.Value = ConvertEntityValue("DisplayName", entityData("DisplayName"))
        cell.Font.Color = RGB(255, 0, 0)
    End If
    
    If entityData.Exists("DisplayCollectionName") Then
        Set cell = wsTable.Range("E6")
        cell.Value = ConvertEntityValue("DisplayCollectionName", entityData("DisplayCollectionName"))
        cell.Font.Color = RGB(255, 0, 0)
    End If
    
    If entityData.Exists("SchemaName") Then
        Set cell = wsTable.Range("E7")
        cell.Value = ConvertEntityValue("SchemaName", entityData("SchemaName"))
        cell.Font.Color = RGB(255, 0, 0)
    End If
    
    If entityData.Exists("Description") Then
        Set cell = wsTable.Range("E8")
        cell.Value = ConvertEntityValue("Description", entityData("Description"))
        cell.Font.Color = RGB(255, 0, 0)
    End If
    
    If entityData.Exists("TableType") Then
        Set cell = wsTable.Range("E9")
        cell.Value = ConvertEntityValue("TableType", entityData("TableType"))
        cell.Font.Color = RGB(255, 0, 0)
    End If
    
    If entityData.Exists("OwnershipType") Then
        Set cell = wsTable.Range("E10")
        cell.Value = ConvertEntityValue("OwnershipType", entityData("OwnershipType"))
        cell.Font.Color = RGB(255, 0, 0)
    End If
    
    If entityData.Exists("PrimaryImageAttribute") Then
        Set cell = wsTable.Range("E11")
        cell.Value = ConvertEntityValue("PrimaryImageAttribute", entityData("PrimaryImageAttribute"))
        cell.Font.Color = RGB(255, 0, 0)
    End If
    
    If entityData.Exists("EntityColor") Then
        Set cell = wsTable.Range("E12")
        cell.Value = ConvertEntityValue("EntityColor", entityData("EntityColor"))
        cell.Font.Color = RGB(255, 0, 0)
    End If
    
    If entityData.Exists("HasNotes") Then
        Set cell = wsTable.Range("E13")
        cell.Value = ConvertEntityValue("HasNotes", entityData("HasNotes"))
        cell.Font.Color = RGB(255, 0, 0)
    End If
    
    If entityData.Exists("IsDuplicateDetectionEnabled") Then
        Set cell = wsTable.Range("E20")
        cell.Value = ConvertEntityValue("IsDuplicateDetectionEnabled", entityData("IsDuplicateDetectionEnabled"))
        cell.Font.Color = RGB(255, 0, 0)
    End If
    
    If entityData.Exists("ChangeTrackingEnabled") Then
        Set cell = wsTable.Range("E21")
        cell.Value = ConvertEntityValue("ChangeTrackingEnabled", entityData("ChangeTrackingEnabled"))
        cell.Font.Color = RGB(255, 0, 0)
    End If
    
    If entityData.Exists("IsKnowledgeManagementEnabled") Then
        Set cell = wsTable.Range("E22")
        cell.Value = ConvertEntityValue("IsKnowledgeManagementEnabled", entityData("IsKnowledgeManagementEnabled"))
        cell.Font.Color = RGB(255, 0, 0)
    End If
    
    If entityData.Exists("EntityHelpUrlEnabled") Then
        Set cell = wsTable.Range("E23")
        cell.Value = ConvertEntityValue("EntityHelpUrlEnabled", entityData("EntityHelpUrlEnabled"))
        cell.Font.Color = RGB(255, 0, 0)
    End If
    
    If entityData.Exists("EntityHelpUrl") Then
        Set cell = wsTable.Range("E24")
        cell.Value = ConvertEntityValue("EntityHelpUrl", entityData("EntityHelpUrl"))
        cell.Font.Color = RGB(255, 0, 0)
    End If
    
    If entityData.Exists("IsAuditEnabled") Then
        Set cell = wsTable.Range("E25")
        cell.Value = ConvertEntityValue("IsAuditEnabled", entityData("IsAuditEnabled"))
        cell.Font.Color = RGB(255, 0, 0)
    End If
    
    If entityData.Exists("IsQuickCreateEnabled") Then
        Set cell = wsTable.Range("E26")
        cell.Value = ConvertEntityValue("IsQuickCreateEnabled", entityData("IsQuickCreateEnabled"))
        cell.Font.Color = RGB(255, 0, 0)
    End If
    
    If entityData.Exists("HasActivities") Then
        Set cell = wsTable.Range("E27")
        cell.Value = ConvertEntityValue("HasActivities", entityData("HasActivities"))
        cell.Font.Color = RGB(255, 0, 0)
    End If
    
    If entityData.Exists("IsMailMergeEnabled") Then
        Set cell = wsTable.Range("E28")
        cell.Value = ConvertEntityValue("IsMailMergeEnabled", entityData("IsMailMergeEnabled"))
        cell.Font.Color = RGB(255, 0, 0)
    End If
    
    If entityData.Exists("IsSLAEnabled") Then
        Set cell = wsTable.Range("E29")
        cell.Value = ConvertEntityValue("IsSLAEnabled", entityData("IsSLAEnabled"))
        cell.Font.Color = RGB(255, 0, 0)
    End If
    
    If entityData.Exists("IsDocumentManagementEnabled") Then
        Set cell = wsTable.Range("E31")
        cell.Value = ConvertEntityValue("IsDocumentManagementEnabled", entityData("IsDocumentManagementEnabled"))
        cell.Font.Color = RGB(255, 0, 0)
    End If
    
    If entityData.Exists("IsConnectionsEnabled") Then
        Set cell = wsTable.Range("E32")
        cell.Value = ConvertEntityValue("IsConnectionsEnabled", entityData("IsConnectionsEnabled"))
        cell.Font.Color = RGB(255, 0, 0)
    End If
    
    If entityData.Exists("AutoCreateAccessTeams") Then
        Set cell = wsTable.Range("E34")
        cell.Value = ConvertEntityValue("AutoCreateAccessTeams", entityData("AutoCreateAccessTeams"))
        cell.Font.Color = RGB(255, 0, 0)
    End If
    
    If entityData.Exists("HasFeedback") Then
        Set cell = wsTable.Range("E35")
        cell.Value = ConvertEntityValue("HasFeedback", entityData("HasFeedback"))
        cell.Font.Color = RGB(255, 0, 0)
    End If
    
    If entityData.Exists("IsValidForAdvancedFind") Then
        Set cell = wsTable.Range("E36")
        cell.Value = ConvertEntityValue("IsValidForAdvancedFind", entityData("IsValidForAdvancedFind"))
        cell.Font.Color = RGB(255, 0, 0)
    End If
    
    If entityData.Exists("IsAvailableOffline") Then
        Set cell = wsTable.Range("E37")
        cell.Value = ConvertEntityValue("IsAvailableOffline", entityData("IsAvailableOffline"))
        cell.Font.Color = RGB(255, 0, 0)
    End If
    
    If entityData.Exists("IsValidForQueue") Then
        Set cell = wsTable.Range("E38")
        cell.Value = ConvertEntityValue("IsValidForQueue", entityData("IsValidForQueue"))
        cell.Font.Color = RGB(255, 0, 0)
    End If

End Sub


'========================================================================
'  attribute.xlsx → テンプレートへ出力
'
'  出力先：シート「テーブル」とシート「フィールド」
'  変更したセルは赤文字にする
'========================================================================
Public Sub SetAttributeInfoToTemplate(wbOut As Workbook, wbAttribute As Workbook)

    Dim wsAttribute As Worksheet
    Dim wsTable As Worksheet
    Dim wsField As Worksheet
    Dim lastCol As Long
    Dim lastRow As Long
    Dim col As Long
    Dim row As Long
    Dim colIndex As Object
    Dim primaryNameRow As Long
    Dim cell As Range
    
    Set wsAttribute = wbAttribute.Sheets(1)
    Set wsTable = wbOut.Sheets("テーブル")
    Set wsField = wbOut.Sheets("フィールド")
    
    '▼ 列名のインデックスを取得
    Set colIndex = CreateObject("Scripting.Dictionary")
    lastCol = wsAttribute.Cells(1, wsAttribute.Columns.Count).End(xlToLeft).Column
    
    For col = 1 To lastCol
        Dim colName As String
        colName = Trim(wsAttribute.Cells(1, col).Value)
        If colName <> "" Then
            colIndex(colName) = col
        End If
    Next col
    
    '▼ データ行数を取得
    lastRow = wsAttribute.Cells(wsAttribute.Rows.Count, 1).End(xlUp).Row
    
    '▼ IsPrimaryName が True の行を探す
    primaryNameRow = 0
    If colIndex.Exists("IsPrimaryName") Then
        Dim primaryNameCol As Long
        primaryNameCol = colIndex("IsPrimaryName")
        For row = 2 To lastRow
            If LCase(Trim(wsAttribute.Cells(row, primaryNameCol).Value)) = "true" Then
                primaryNameRow = row
                Exit For
            End If
        Next row
    End If
    
    '▼ シート「テーブル」に IsPrimaryName が True の行のデータを設定
    If primaryNameRow > 0 Then
        If colIndex.Exists("DisplayName") Then
            Set cell = wsTable.Range("E14")
            cell.Value = Trim(wsAttribute.Cells(primaryNameRow, colIndex("DisplayName")).Value)
            cell.Font.Color = RGB(255, 0, 0)
        End If
        
        If colIndex.Exists("Description") Then
            Set cell = wsTable.Range("E15")
            cell.Value = Trim(wsAttribute.Cells(primaryNameRow, colIndex("Description")).Value)
            cell.Font.Color = RGB(255, 0, 0)
        End If
        
        If colIndex.Exists("SchemaName") Then
            Set cell = wsTable.Range("E16")
            cell.Value = Trim(wsAttribute.Cells(primaryNameRow, colIndex("SchemaName")).Value)
            cell.Font.Color = RGB(255, 0, 0)
        End If
        
        If colIndex.Exists("LogicalName") Then
            Set cell = wsTable.Range("E17")
            cell.Value = Trim(wsAttribute.Cells(primaryNameRow, colIndex("LogicalName")).Value)
            cell.Font.Color = RGB(255, 0, 0)
        End If
        
        If colIndex.Exists("RequiredLevel") Then
            Set cell = wsTable.Range("E18")
            cell.Value = Trim(wsAttribute.Cells(primaryNameRow, colIndex("RequiredLevel")).Value)
            cell.Font.Color = RGB(255, 0, 0)
        End If
    End If
    
    '▼ シート「フィールド」に全 attribute データを記述（7行目から）
    Dim fieldRow As Long
    fieldRow = 7
    
    For row = 2 To lastRow
        'E列: SchemaName
        If colIndex.Exists("SchemaName") Then
            Set cell = wsField.Cells(fieldRow, 5)  'E列
            cell.Value = Trim(wsAttribute.Cells(row, colIndex("SchemaName")).Value)
            cell.Font.Color = RGB(255, 0, 0)
        End If
        
        'F列: DisplayName
        If colIndex.Exists("DisplayName") Then
            Set cell = wsField.Cells(fieldRow, 6)  'F列
            cell.Value = Trim(wsAttribute.Cells(row, colIndex("DisplayName")).Value)
            cell.Font.Color = RGB(255, 0, 0)
        End If
        
        'G列: LogicalName
        If colIndex.Exists("LogicalName") Then
            Set cell = wsField.Cells(fieldRow, 7)  'G列
            cell.Value = Trim(wsAttribute.Cells(row, colIndex("LogicalName")).Value)
            cell.Font.Color = RGB(255, 0, 0)
        End If
        
        'H列: IsCustomAttribute
        If colIndex.Exists("IsCustomAttribute") Then
            Set cell = wsField.Cells(fieldRow, 8)  'H列
            cell.Value = ConvertAttributeValue("IsCustomAttribute", Trim(wsAttribute.Cells(row, colIndex("IsCustomAttribute")).Value))
            cell.Font.Color = RGB(255, 0, 0)
        End If
        
        'I列: IsPrimaryID
        If colIndex.Exists("IsPrimaryID") Then
            Set cell = wsField.Cells(fieldRow, 9)  'I列
            cell.Value = ConvertAttributeValue("IsPrimaryID", Trim(wsAttribute.Cells(row, colIndex("IsPrimaryID")).Value))
            cell.Font.Color = RGB(255, 0, 0)
        End If
        
        'J列: IsPrimaryName
        If colIndex.Exists("IsPrimaryName") Then
            Set cell = wsField.Cells(fieldRow, 10)  'J列
            cell.Value = ConvertAttributeValue("IsPrimaryName", Trim(wsAttribute.Cells(row, colIndex("IsPrimaryName")).Value))
            cell.Font.Color = RGB(255, 0, 0)
        End If
        
        'K列: AttributeTypeName
        If colIndex.Exists("AttributeTypeName") Then
            Set cell = wsField.Cells(fieldRow, 11)  'K列
            cell.Value = Trim(wsAttribute.Cells(row, colIndex("AttributeTypeName")).Value)
            cell.Font.Color = RGB(255, 0, 0)
        End If
        
        'M列: RequiredLevel
        If colIndex.Exists("RequiredLevel") Then
            Set cell = wsField.Cells(fieldRow, 13)  'M列
            cell.Value = Trim(wsAttribute.Cells(row, colIndex("RequiredLevel")).Value)
            cell.Font.Color = RGB(255, 0, 0)
        End If
        
        'Y列: IsAuditEnabled
        If colIndex.Exists("IsAuditEnabled") Then
            Set cell = wsField.Cells(fieldRow, 25)  'Y列
            cell.Value = ConvertAttributeValue("IsAuditEnabled", Trim(wsAttribute.Cells(row, colIndex("IsAuditEnabled")).Value))
            cell.Font.Color = RGB(255, 0, 0)
        End If
        
        'AA列: IsGlobalFilterEnabled
        If colIndex.Exists("IsGlobalFilterEnabled") Then
            Set cell = wsField.Cells(fieldRow, 27)  'AA列
            cell.Value = ConvertAttributeValue("IsGlobalFilterEnabled", Trim(wsAttribute.Cells(row, colIndex("IsGlobalFilterEnabled")).Value))
            cell.Font.Color = RGB(255, 0, 0)
        End If
        
        'AB列: IsSortableEnabled
        If colIndex.Exists("IsSortableEnabled") Then
            Set cell = wsField.Cells(fieldRow, 28)  'AB列
            cell.Value = ConvertAttributeValue("IsSortableEnabled", Trim(wsAttribute.Cells(row, colIndex("IsSortableEnabled")).Value))
            cell.Font.Color = RGB(255, 0, 0)
        End If
        
        'AC列: IsSearchable
        If colIndex.Exists("IsSearchable") Then
            Set cell = wsField.Cells(fieldRow, 29)  'AC列
            cell.Value = ConvertAttributeValue("IsSearchable", Trim(wsAttribute.Cells(row, colIndex("IsSearchable")).Value))
            cell.Font.Color = RGB(255, 0, 0)
        End If
        
        'AF列: Description
        If colIndex.Exists("Description") Then
            Set cell = wsField.Cells(fieldRow, 32)  'AF列
            cell.Value = Trim(wsAttribute.Cells(row, colIndex("Description")).Value)
            cell.Font.Color = RGB(255, 0, 0)
        End If
        
        'AG列: IsSecret
        If colIndex.Exists("IsSecret") Then
            Set cell = wsField.Cells(fieldRow, 33)  'AG列
            cell.Value = ConvertAttributeValue("IsSecret", Trim(wsAttribute.Cells(row, colIndex("IsSecret")).Value))
            cell.Font.Color = RGB(255, 0, 0)
        End If
        
        fieldRow = fieldRow + 1
    Next row

End Sub


'========================================================================
'  attribute 値変換（True/False）
'========================================================================
Private Function ConvertAttributeValue(key As String, val As String) As String
    
    val = Trim(val)
    
    If LCase(val) = "true" Then
        ConvertAttributeValue = "ON"
    ElseIf LCase(val) = "false" Then
        ConvertAttributeValue = "-"
    Else
        ConvertAttributeValue = val
    End If

End Function
