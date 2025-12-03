Option Explicit

'────────────────────────────────────────
'  メイン処理（attribute 不使用バージョン）
'────────────────────────────────────────
Public Sub メイン処理_Entityのみ()

    Dim wbThis As Workbook
    Dim folderBase As String
    Dim folderEntity As String
    Dim folderTemplate As String
    Dim folderOutput As String
    
    Dim templatePath As String
    Dim entityFile As String
    Dim entityPath As String
    Dim outputPath As String
    
    Dim wbTemplate As Workbook
    Dim wbOut As Workbook
    Dim wbEntity As Workbook
    
    On Error GoTo ERR_HANDLER
    
    Set wbThis = ThisWorkbook
    folderBase = wbThis.Path & "\"
    
    '▼ 必要なフォルダ
    folderEntity = folderBase & "10_entity\"
    folderTemplate = folderBase & "template\"
    folderOutput = folderBase & "30_create_file\"
    
    '▼ フォルダ存在チェック
    If Dir(folderEntity, vbDirectory) = "" Then Err.Raise 100, , "10_entity フォルダがありません。"
    If Dir(folderTemplate, vbDirectory) = "" Then Err.Raise 101, , "template フォルダがありません。"
    If Dir(folderOutput, vbDirectory) = "" Then Err.Raise 102, , "30_create_file フォルダがありません。"
    
    templatePath = folderTemplate & "template.xlsx"
    If Dir(templatePath) = "" Then Err.Raise 103, , "template.xlsx が見つかりません。"
    
    '▼ entity フォルダの全Excelを処理
    entityFile = Dir(folderEntity & "*.xlsx")
    If entityFile = "" Then Err.Raise 104, , "10_entity に処理対象ファイルがありません。"
    
    Do While entityFile <> ""
    
        entityPath = folderEntity & entityFile
        
        '▼ entity を開く
        Set wbEntity = Workbooks.Open(entityPath, ReadOnly:=True)
        
        '▼ テンプレートを開く
        Set wbTemplate = Workbooks.Open(templatePath, ReadOnly:=True)
        
        '▼ 出力先ファイル名
        outputPath = folderOutput & entityFile
        
        '▼ テンプレートをコピー
        wbTemplate.SaveCopyAs outputPath
        
        'コピーしたものを開く
        Set wbOut = Workbooks.Open(outputPath)
        
        '=====================================
        '  ★ entity の情報だけを出力する処理
        '=====================================
        Call SetEntityInfoToTemplate(wbOut, wbEntity)
        
        '▼ 正常終了時のクローズ
        wbOut.Close SaveChanges:=True
        wbTemplate.Close SaveChanges:=False
        wbEntity.Close SaveChanges:=False
        
        entityFile = Dir()
    Loop
    
    MsgBox "entity データのみの出力が完了しました。", vbInformation
    Exit Sub

'────────────────────────────────────────
ERR_HANDLER:
'────────────────────────────────────────
    On Error Resume Next
    If Not wbOut Is Nothing Then wbOut.Close SaveChanges:=False
    If Not wbTemplate Is Nothing Then wbTemplate.Close SaveChanges:=False
    If Not wbEntity Is Nothing Then wbEntity.Close SaveChanges:=False
    
    MsgBox "エラー：" & Err.Description, vbCritical

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
                ConvertEntityValue = "チェック"
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
'  出力先：テンプレートの "Entity" シート
'  出力形式：A列＝日本語名、B列＝値
'========================================================================
Public Sub SetEntityInfoToTemplate(wbOut As Workbook, wbEntity As Workbook)

    Dim wsEntity As Worksheet
    Dim wsOut As Worksheet
    Dim dic As Object
    Dim lastCol As Long
    Dim col As Long
    Dim engKey As String, val As String, jpKey As String
    Dim rowOut As Long: rowOut = 2
    
    Set wsEntity = wbEntity.Sheets(1)
    Set wsOut = wbOut.Sheets("Entity")   'テンプレートのシート名
    Set dic = GetEntityMappingDict()
    
    lastCol = wsEntity.Cells(1, wsEntity.Columns.Count).End(xlToLeft).Column
    
    For col = 1 To lastCol
        
        engKey = Trim(wsEntity.Cells(1, col).Value)
        val = Trim(wsEntity.Cells(2, col).Value)
        
        If dic.Exists(engKey) Then
            
            jpKey = dic(engKey)
            
            wsOut.Cells(rowOut, 1).Value = jpKey
            wsOut.Cells(rowOut, 2).Value = ConvertEntityValue(engKey, val)
            
            rowOut = rowOut + 1
        End If
        
    Next col

End Sub
