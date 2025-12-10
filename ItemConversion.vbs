Option Explicit

'────────────────────────────────────────
'  処理：ドラッグ&ドロップでフォルダを受け取り、
'  フォルダ内のExcelファイルを順次処理し、
'  ・シート「テーブル」のE12の値を変換
'    OrganizationOwned → 組織
'    UserOwned → ユーザーまたはチーム
'  ・シート「テーブル」のE11の値を変換
'    Standard → 標準
'    activity → 活動
'    virtual → 仮想
'  ・シート「テーブル」のE20の値を変換
'    None → 任意
'    ApplicationRequired → 必須
'  ※エラーハンドリングと処理のベースは
'    CorrectionToDefinitionDocument.vbsを参考にしています
'────────────────────────────────────────

Dim fso, excel, wb, wsTable, ws
Dim folderPath, folder, file
Dim fileName, filePath, fileExt
Dim cellValue, convertedValue
Dim linkSources, linkIndex, linkCount

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
    
    ' Excelファイルの拡張子をチェック
    If fileExt = "xlsx" Or fileExt = "xls" Or fileExt = "xlsm" Or fileExt = "xlsb" Then
        filePath = file.Path
        
        ' 前のループで作成されたオブジェクトを解放（メモリリーク防止）
        On Error Resume Next
        If Not wsTable Is Nothing Then
            Set wsTable = Nothing
        End If
        If Not ws Is Nothing Then
            Set ws = Nothing
        End If
        On Error GoTo 0
        
        On Error Resume Next
        ' Excelファイルを開く（リンクの更新を無効にする）
        Set wb = excel.Workbooks.Open(filePath, 0, False)
        
        If Err.Number = 0 Then
            On Error GoTo 0
            ' リンクの更新を無効にする（警告を表示しない）
            On Error Resume Next
            wb.UpdateLinks = 0  ' xlUpdateLinksNever
            If Err.Number <> 0 Then
                Err.Clear
            End If
            On Error GoTo 0
            
            ' シート「テーブル」を取得
            On Error Resume Next
            Set wsTable = Nothing
            
            ' シート名で検索
            For Each ws In wb.Sheets
                If ws.Name = "テーブル" Then
                    Set wsTable = ws
                    Exit For
                End If
            Next
            
            If Err.Number <> 0 Then
                Err.Clear
            End If
            On Error GoTo 0
            
            ' ▼ シート「テーブル」の各セルの値を変換
            If Not wsTable Is Nothing Then
                ' ▼ E12の値を変換（OrganizationOwned → 組織、UserOwned → ユーザーまたはチーム）
                On Error Resume Next
                cellValue = wsTable.Cells(12, 5).Value  ' E12 = 5列目、12行目
                If Err.Number = 0 And Not IsEmpty(cellValue) Then
                    cellValue = Trim(CStr(cellValue))
                    convertedValue = ""
                    
                    ' 大文字小文字を区別せずに比較
                    If StrComp(cellValue, "OrganizationOwned", vbTextCompare) = 0 Then
                        convertedValue = "組織"
                    ElseIf StrComp(cellValue, "UserOwned", vbTextCompare) = 0 Then
                        convertedValue = "ユーザーまたはチーム"
                    End If
                    
                    ' 変換値が設定された場合のみ更新
                    If convertedValue <> "" Then
                        wsTable.Cells(12, 5).Value = convertedValue
                    End If
                End If
                Err.Clear
                On Error GoTo 0
                
                ' ▼ E11の値を変換（Standard → 標準、activity → 活動、virtual → 仮想）
                On Error Resume Next
                cellValue = wsTable.Cells(11, 5).Value  ' E11 = 5列目、11行目
                If Err.Number = 0 And Not IsEmpty(cellValue) Then
                    cellValue = Trim(CStr(cellValue))
                    convertedValue = ""
                    
                    ' 大文字小文字を区別せずに比較
                    If StrComp(cellValue, "Standard", vbTextCompare) = 0 Then
                        convertedValue = "標準"
                    ElseIf StrComp(cellValue, "activity", vbTextCompare) = 0 Then
                        convertedValue = "活動"
                    ElseIf StrComp(cellValue, "virtual", vbTextCompare) = 0 Then
                        convertedValue = "仮想"
                    End If
                    
                    ' 変換値が設定された場合のみ更新
                    If convertedValue <> "" Then
                        wsTable.Cells(11, 5).Value = convertedValue
                    End If
                End If
                Err.Clear
                On Error GoTo 0
                
                ' ▼ E20の値を変換（None → 任意、ApplicationRequired → 必須）
                On Error Resume Next
                cellValue = wsTable.Cells(20, 5).Value  ' E20 = 5列目、20行目
                If Err.Number = 0 And Not IsEmpty(cellValue) Then
                    cellValue = Trim(CStr(cellValue))
                    convertedValue = ""
                    
                    ' 大文字小文字を区別せずに比較
                    If StrComp(cellValue, "None", vbTextCompare) = 0 Then
                        convertedValue = "任意"
                    ElseIf StrComp(cellValue, "ApplicationRequired", vbTextCompare) = 0 Then
                        convertedValue = "必須"
                    End If
                    
                    ' 変換値が設定された場合のみ更新
                    If convertedValue <> "" Then
                        wsTable.Cells(20, 5).Value = convertedValue
                    End If
                End If
                Err.Clear
                On Error GoTo 0
            End If
            
            ' ▼ 「保存しますか？」ダイアログを防ぐため、保存前に計算を実行
            ' 計算モードを自動に戻してから計算を実行
            ' 注意：循環参照がある場合は計算が終わらない可能性がある
            ' 高速化のため、計算モード変更と計算を最後に一度だけ実行
            On Error Resume Next
            excel.Calculation = -4105   ' xlCalculationAutomatic
            ' 計算を実行（循環参照がある場合は時間がかかる可能性がある）
            ' Calculateは同期的に実行されるため、計算が完了するまで待機する
            wb.Calculate   ' ブック全体を計算
            If Err.Number <> 0 Then
                ' 計算エラーが発生した場合（循環参照など）は計算モードを手動に戻す
                excel.Calculation = -4135   ' xlCalculationManual
                Err.Clear
            End If
            
            ' ▼ 外部リンクを削除して警告を防ぐ
            On Error Resume Next
            linkSources = wb.LinkSources(1)  ' xlExcelLinks
            If Err.Number = 0 Then
                ' リンクが存在する場合
                If IsArray(linkSources) Then
                    On Error Resume Next
                    linkCount = UBound(linkSources) + 1
                    If Err.Number = 0 And linkCount > 0 Then
                        For linkIndex = 0 To linkCount - 1
                            On Error Resume Next
                            wb.BreakLink linkSources(linkIndex), 1  ' xlLinkTypeExcelLinks
                            Err.Clear
                            On Error GoTo 0
                        Next
                    End If
                    Err.Clear
                ElseIf Not IsEmpty(linkSources) And linkSources <> "" Then
                    ' リンクが1つの場合（配列ではなく文字列）
                    On Error Resume Next
                    wb.BreakLink linkSources, 1  ' xlLinkTypeExcelLinks
                    Err.Clear
                End If
            End If
            Err.Clear
            On Error GoTo 0
            
            ' 保存ダイアログを防ぐため、SavedプロパティをTrueに設定
            wb.Saved = True
            ' リンクの更新を無効にする（保存時の警告を防ぐ）
            On Error Resume Next
            wb.UpdateLinks = 0  ' xlUpdateLinksNever
            If Err.Number <> 0 Then
                Err.Clear
            End If
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
