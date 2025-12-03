Option Explicit

Dim fso, excel, wbSrc, wbDst, templatePath, srcPath, outPath
Dim mapping, headerDict
Dim srcSheet, dstSheetCover, dstSheetTable
Dim lastCol, i, colName, value, beforeValue

Set fso = CreateObject("Scripting.FileSystemObject")

' ▼ パス設定
templatePath = fso.GetAbsolutePathName("テンプレート.xlsx")
srcPath      = fso.GetAbsolutePathName("実行対象エクセル\エンティティA.xlsx")

' ▼ Excel 起動
Set excel = CreateObject("Excel.Application")
excel.Visible = False

' ▼ 元データ読み込み
Set wbSrc = excel.Workbooks.Open(srcPath)
Set srcSheet = wbSrc.Sheets(1)

'=========================
' ヘッダー列の位置取得
'=========================
Set headerDict = CreateObject("Scripting.Dictionary")

lastCol = srcSheet.Cells(1, srcSheet.Columns.Count).End(-4159).Column 'xlToLeft

For i = 1 To lastCol
    colName = Trim(CStr(srcSheet.Cells(1, i).Value))
    If colName <> "" Then headerDict(colName) = i
Next

'=========================
' LogicalName / DisplayName → 出力名生成
'=========================
Dim logicalName, displayName, outFileName

logicalName = srcSheet.Cells(2, headerDict("LogicalName")).Value
displayName = srcSheet.Cells(2, headerDict("DisplayName")).Value

outFileName = "エンティティ定義書_" & logicalName & "_" & displayName & "_v0.0.xlsx"
outPath = fso.GetAbsolutePathName("作成済エクセル\" & outFileName)

'=========================
' テンプレートをコピー
'=========================
' シートをインデックスで取得（日本語文字列を避ける）
Const SHEET_INDEX_COVER = 1  ' 表紙シート
Const SHEET_INDEX_TABLE = 2  ' テーブルシート

Set wbDst = excel.Workbooks.Open(templatePath)
wbDst.SaveAs outPath

Set dstSheetCover = wbDst.Sheets(SHEET_INDEX_COVER)
Set dstSheetTable = wbDst.Sheets(SHEET_INDEX_TABLE)

'=========================
' マッピング（EntityMetadata → セル）
'=========================
Set mapping = CreateObject("Scripting.Dictionary")

' --- 表紙シート（インデックス1） ---
mapping.Add "DisplayName|1", "W21"

' --- テーブルシート（インデックス2） ---
mapping.Add "DisplayName|2", "E5"
mapping.Add "DisplayCollectionName|2", "E6"
mapping.Add "SchemaName|2", "E7"
mapping.Add "Description|2", "E8"
mapping.Add "TableType|2", "E9"
mapping.Add "OwnershipType|2", "E10"
mapping.Add "PrimaryImageAttribute|2", "E11"
mapping.Add "EntityColor|2", "E12"

mapping.Add "IsDuplicateDetectionEnabled|2", "E20"
mapping.Add "ChangeTrackingEnabled|2", "E21"
mapping.Add "IsKnowledgeManagementEnabled|2", "E22"
mapping.Add "EntityHelpUrlEnabled|2", "E23"
mapping.Add "EntityHelpUrl|2", "E24"
mapping.Add "IsAuditEnabled|2", "E25"
mapping.Add "IsQuickCreateEnabled|2", "E26"
mapping.Add "HasActivities|2", "E27"
mapping.Add "IsMailMergeEnabled|2", "E28"
mapping.Add "IsSLAEnabled|2", "E29"
mapping.Add "IsDocumentManagementEnabled|2", "E31"
mapping.Add "IsConnectionsEnabled|2", "E32"
mapping.Add "AutoCreateAccessTeams|2", "E34"
mapping.Add "HasFeedback|2", "E35"
mapping.Add "IsAvailableOffline|2", "E37"
mapping.Add "IsValidForQueue|2", "E38"

'=========================
' 変換関数（True/False → 記号, TableType → 日本語）
'=========================
Function ConvertValue(metaName, raw)
    Dim lowerVal
    lowerVal = LCase(CStr(raw))

    '--- TableType の変換 ---
    If metaName = "TableType" Then
        Select Case lowerVal
            Case "standard"
                ConvertValue = "標準"
                Exit Function
            Case "activity"
                ConvertValue = "活動"
                Exit Function
            Case "virtual"
                ConvertValue = "仮想"
                Exit Function
        End Select
    End If

    '--- True / False の変換 ---
    If lowerVal = "true" Then
        ConvertValue = "✔️"
    ElseIf lowerVal = "false" Or lowerVal = "" Then
        ConvertValue = "－"
    Else
        ConvertValue = raw
    End If
End Function

'=========================
' 転記処理（赤字判定あり）
'=========================
Dim key, parts, metaName, sheetIndex, cellAddr, dstSheet

For Each key In mapping.Keys

    parts = Split(key, "|")
    metaName = parts(0)
    sheetIndex = CInt(parts(1))  ' 数値インデックスに変換
    cellAddr = mapping(key)

    ' 元データ取得
    If headerDict.Exists(metaName) Then
        value = srcSheet.Cells(2, headerDict(metaName)).Value
    Else
        value = "(Not Found)"
    End If

    ' 値変換を適用
    value = ConvertValue(metaName, value)

    ' 出力先シート（インデックスで判定）
    If sheetIndex = SHEET_INDEX_COVER Then
        Set dstSheet = dstSheetCover
    ElseIf sheetIndex = SHEET_INDEX_TABLE Then
        Set dstSheet = dstSheetTable
    Else
        ' エラー処理（必要に応じて）
        WScript.Echo "警告: 不明なシートインデックス " & sheetIndex
        Set dstSheet = dstSheetTable
    End If

    beforeValue = dstSheet.Range(cellAddr).Value
    dstSheet.Range(cellAddr).Value = value

    ' 変更があれば赤字
    If CStr(beforeValue) <> CStr(value) Then
        dstSheet.Range(cellAddr).Font.Color = RGB(255, 0, 0)
    End If
Next

'=========================
' 終了処理
'=========================
wbDst.Save
wbDst.Close False
wbSrc.Close False
excel.Quit

WScript.Echo "Complete -> " & outPath
