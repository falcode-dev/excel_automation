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
Set wbDst = excel.Workbooks.Open(templatePath)
wbDst.SaveAs outPath

Set dstSheetCover = wbDst.Sheets("表紙")
Set dstSheetTable = wbDst.Sheets("テーブル")

'=========================
' マッピング（EntityMetadata → セル）
'=========================
Set mapping = CreateObject("Scripting.Dictionary")

' --- 表紙 ---
mapping.Add "DisplayName|表紙", "W21"

' --- テーブル ---
mapping.Add "DisplayName|テーブル", "E5"
mapping.Add "DisplayCollectionName|テーブル", "E6"
mapping.Add "SchemaName|テーブル", "E7"
mapping.Add "Description|テーブル", "E8"
mapping.Add "TableType|テーブル", "E9"
mapping.Add "OwnershipType|テーブル", "E10"
mapping.Add "PrimaryImageAttribute|テーブル", "E11"
mapping.Add "EntityColor|テーブル", "E12"

mapping.Add "IsDuplicateDetectionEnabled|テーブル", "E20"
mapping.Add "ChangeTrackingEnabled|テーブル", "E21"
mapping.Add "IsKnowledgeManagementEnabled|テーブル", "E22"
mapping.Add "EntityHelpUrlEnabled|テーブル", "E23"
mapping.Add "EntityHelpUrl|テーブル", "E24"
mapping.Add "IsAuditEnabled|テーブル", "E25"
mapping.Add "IsQuickCreateEnabled|テーブル", "E26"
mapping.Add "HasActivities|テーブル", "E27"
mapping.Add "IsMailMergeEnabled|テーブル", "E28"
mapping.Add "IsSLAEnabled|テーブル", "E29"
mapping.Add "IsDocumentManagementEnabled|テーブル", "E31"
mapping.Add "IsConnectionsEnabled|テーブル", "E32"
mapping.Add "AutoCreateAccessTeams|テーブル", "E34"
mapping.Add "HasFeedback|テーブル", "E35"
mapping.Add "IsAvailableOffline|テーブル", "E37"
mapping.Add "IsValidForQueue|テーブル", "E38"

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
Dim key, parts, metaName, sheetName, cellAddr, dstSheet

For Each key In mapping.Keys

    parts = Split(key, "|")
    metaName = parts(0)
    sheetName = parts(1)
    cellAddr = mapping(key)

    ' 元データ取得
    If headerDict.Exists(metaName) Then
        value = srcSheet.Cells(2, headerDict(metaName)).Value
    Else
        value = "(Not Found)"
    End If

    ' 値変換を適用
    value = ConvertValue(metaName, value)

    ' 出力先シート
    If sheetName = "表紙" Then
        Set dstSheet = dstSheetCover
    Else
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

WScript.Echo "完了 → " & outPath
