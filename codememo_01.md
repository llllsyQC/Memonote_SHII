' ================================
' 📦 Module: modLogUtility
' 日志出力 + エラー回復支援
' ================================

Public gLogFilePath As String
Public gHasErrorOccurred As Boolean

' 📌 初期化：ログファイル準備
Sub InitLogger(yyyymm As String)
    Dim folderPath As String
    folderPath = "\\bbwcfs.local\share4\SBM1\SharePrj\大宮データセンター運用フォルダ\14.定例会\編集用資料\工数取得マクロ作成済\ログ\" & yyyymm
    
    MkDirRecursive folderPath
    
    gLogFilePath = folderPath & "\log_" & Format(Now, "yyyymmdd_HHmmss") & ".txt"
    
    Open gLogFilePath For Output As #1
    Print #1, "【ログ開始】" & Now
    Close #1
End Sub

' 📌 ログ追記
Sub WriteLog(msg As String)
    On Error Resume Next
    Open gLogFilePath For Append As #1
    Print #1, "[" & Format(Now, "yyyy/mm/dd HH:MM:SS") & "] " & msg
    Close #1
End Sub

' 📌 フォルダ存在チェック + 再帰作成
Sub MkDirRecursive(ByVal path As String)
    Dim tempPath As String, folders() As String
    Dim i As Integer
    folders = Split(path, "\")
    tempPath = folders(0)
    For i = 1 To UBound(folders)
        tempPath = tempPath & "\" & folders(i)
        If Dir(tempPath, vbDirectory) = "" Then MkDir tempPath
    Next i
End Sub

' 📌 グローバルロールバック処理（例: エラーハンドラ）
Sub ErrorRollback(Optional ByVal errMsg As String = "")
    gHasErrorOccurred = True
    If errMsg <> "" Then WriteLog "❌ エラー発生：" & errMsg
    MsgBox "エラーが発生しました。" & vbCrLf & errMsg & vbCrLf & "ログを確認してください。", vbCritical, "処理中断"
End Sub

' ================================
' 📦 Module: modMain
' 工数集計_自動処理の起点
' ================================

Sub 工数集計_自動処理()
    On Error GoTo ERR_HANDLER
    Dim 年 As String, 月 As String, yyyymm As String
    Dim pathDaily As String
    
    年 = InputBox("年を入力してください（例：2025）")
    If 年 = "" Then MsgBox "年の入力が必要です。": Exit Sub
    If Not IsNumeric(年) Or Len(年) <> 4 Then MsgBox "4桁の西暦で入力してください。": Exit Sub
    
    月 = InputBox("月を入力してください（例：04）")
    If 月 = "" Then MsgBox "月の入力が必要です。": Exit Sub
    If Not IsNumeric(月) Or Val(月) < 1 Or Val(月) > 12 Then MsgBox "1～12の数値を入力してください。": Exit Sub
    If Len(月) = 1 Then 月 = "0" & 月
    
    yyyymm = 年 & 月
    
    ' === ログ初期化 ===
    Call InitLogger(yyyymm)
    Call WriteLog("📌 工数集計 開始（年月：" & yyyymm & "）")
    
    ' === フォルダ生成 ===
    pathDaily = GetTargetFolderPath(年, 月)
    If Dir(pathDaily, vbDirectory) = "" Then
        Call ErrorRollback("フォルダが見つかりません：" & pathDaily)
        Exit Sub
    End If
    Call WriteLog("✅ 対象フォルダ存在確認：" & pathDaily)
    
    ' === メイン処理 ===
    Call Import_AllSources(yyyymm, 年, 月, pathDaily)
    Call Copy_WorkHours_By工数番号(yyyymm, 年, 月)
    Call Process_Ver52(yyyymm, 年, 月)
    
    Call WriteLog("🎉 工数集計 完了：" & yyyymm)
    MsgBox "工数集計が完了しました！", vbInformation

    Exit Sub

ERR_HANDLER:
    Call ErrorRollback("予期しないエラー：" & Err.Description)
End Sub

Function GetTargetFolderPath(ByVal 年 As String, ByVal 月 As String) As String
    Dim basePath As String
    basePath = "\\bbwcfs.local\share4\SBM1\SharePrj\大宮データセンター運用フォルダ\06.日報\06.50定時_工数集計\202504工数ルール改定\マクロ修正中\過去分"
    GetTargetFolderPath = basePath & "\" & 年 & "年\" & 月 & "月"
End Function

' ================================
' 📦 Module: modImport
' データ読み込みと工数転記処理
' ================================

Sub Import_AllSources(yyyymm As String, 年 As String, 月 As String, pathDaily As String)
    On Error GoTo ERR_HANDLER
    Dim fileName1 As String, fileName2 As String, fileName4 As String
    Dim wb1 As Workbook, wb2 As Workbook, wb4 As Workbook
    
    ' === ① 定時作業チェックシート読み込み ===
    fileName1 = Dir(pathDaily & "\①工数集計_定時作業チェックシート_*" & yyyymm & "*.xlsx")
    If fileName1 <> "" Then
        Set wb1 = Workbooks.Open(pathDaily & "\" & fileName1)
        With wb1.Worksheets("定時作業工数詳細")
            .Columns.Hidden = False
            .Rows.Hidden = False
            .Range("M4:AQ57").Copy
            ThisWorkbook.Worksheets("定(日)").Range("B3").PasteSpecial xlPasteValues
            .Range("M59:AQ112").Copy
            ThisWorkbook.Worksheets("定(夜)").Range("B3").PasteSpecial xlPasteValues
        End With
        wb1.Close SaveChanges:=False
        WriteLog "✅ ①定時作業チェックシート読込成功：" & fileName1
    Else
        WriteLog "⚠️ ①定時作業チェックシートが見つかりません"
    End If

    ' === ② 定時外シート読み込み ===
    fileName2 = Dir(pathDaily & "\②工数集計_定時外_*" & yyyymm & "*.xlsx")
    If fileName2 <> "" Then
        Set wb2 = Workbooks.Open(pathDaily & "\" & fileName2)
        wb2.Worksheets("②集計シード").Range("F4:BO57").Copy
        ThisWorkbook.Worksheets("定時外").Range("F7").PasteSpecial xlPasteValues
        wb2.Close False
        WriteLog "✅ ②定時外読込成功：" & fileName2
    Else
        Call ErrorRollback("②定時外ファイルが見つかりません。")
        Exit Sub
    End If

    ' === ④ 日報シート読み込み ===
    fileName4 = Dir(pathDaily & "\④工数集計_日報_*" & yyyymm & "*.xlsx")
    If fileName4 <> "" Then
        Set wb4 = Workbooks.Open(pathDaily & "\" & fileName4)
        wb4.Worksheets("日報").Range("F7:LC60").Copy
        ThisWorkbook.Worksheets("日報").Range("F7").PasteSpecial xlPasteValues
        wb4.Close False
        WriteLog "✅ ④日報読込成功：" & fileName4
    Else
        Call ErrorRollback("④日報ファイルが見つかりません。")
        Exit Sub
    End If

    ' === 年月記入 ===
    With ThisWorkbook.Worksheets("工数取得-都度対応項目（時間）")
        .Range("I5").Value = 年
        .Range("I8").Value = 月
    End With
    WriteLog "✅ 年月情報記入完了"

    Exit Sub
ERR_HANDLER:
    Call ErrorRollback("Import_AllSourcesでエラー：" & Err.Description)
End Sub

Sub Copy_WorkHours_By工数番号(yyyymm As String, 年 As String, 月 As String)
    On Error GoTo ERR_HANDLER
    Dim wsSrc As Worksheet, wsDest As Worksheet
    Dim 工数Map As Object
    Dim i As Integer, srcRow As Long, destRow As Long
    Dim 工数番号 As String
    Dim dayBaseCol As Long, colOffset As Long
    Dim 日勤列 As Long, 夜勤列 As Long
    Dim 日付判定セル As Range

    Set wsSrc = ThisWorkbook.Worksheets("0消し（最終）")
    Set wsDest = ThisWorkbook.Worksheets("工数取得-都度対応項目（時間）")
    Set 工数Map = CreateObject("Scripting.Dictionary")
    
    ' === 工数番号 → 行マッピング作成 ===
    For destRow = 11 To 64
        工数番号 = Trim(CStr(wsDest.Cells(destRow, "I").Value))
        If 工数番号 <> "" Then 工数Map(工数番号) = destRow
    Next

    ' === 1日～31日のループ処理 ===
    For i = 0 To 30
        Set 日付判定セル = wsDest.Cells(11, 4 * i + 15)
        If 日付判定セル.DisplayFormat.Interior.ColorIndex = xlNone Then
            colOffset = 2 ' 休日 → R/S
        Else
            colOffset = 0 ' 平日 → P/Q
        End If
        dayBaseCol = 4 * i + 15 + colOffset

        For srcRow = 6 To 59
            工数番号 = Trim(CStr(wsSrc.Cells(srcRow, "E").Value))
            If 工数Map.exists(工数番号) Then
                destRow = 工数Map(工数番号)
                日勤列 = 2 * i + 6
                夜勤列 = 日勤列 + 1

                With wsDest.Cells(destRow, dayBaseCol)
                    .NumberFormat = "h:mm"
                    .Value = wsSrc.Cells(srcRow, 日勤列).Value
                End With
                With wsDest.Cells(destRow, dayBaseCol + 1)
                    .NumberFormat = "h:mm"
                    .Value = wsSrc.Cells(srcRow, 夜勤列).Value
                End With
            End If
        Next srcRow
    Next i

    WriteLog "✅ 工数番号転記完了"
    Exit Sub
ERR_HANDLER:
    Call ErrorRollback("Copy_WorkHours_By工数番号でエラー：" & Err.Description)
End Sub

' ================================
' 📦 Module: modProcessVer52
' 工数取得ver.5.2（大宮）への転記処理
' ================================

Sub Process_Ver52(yyyymm As String, 年 As String, 月 As String)
    On Error GoTo ERR_HANDLER

    Dim wbVer52 As Workbook, wbSource As Workbook
    Dim wsTarget As Worksheet, wsSource As Worksheet
    Dim pathVer52 As String, pathTemplate As String
    Dim fileNameSource As String, newFileName As String
    Dim yy As String

    ' === パス定義（ログで使われる）===
    pathVer52 = "\\bbwcfs.local\share4\SBM1\SharePrj\大宮データセンター運用フォルダ\14.定例会\編集用資料\【原紙】IDCF工数提出用マクロファイル\"
    pathTemplate = "\\bbwcfs.local\share4\SBM1\SharePrj\大宮データセンター運用フォルダ\06.日報\06.50定時_工数集計\202504工数ルール改定\マクロ修正中\作成済\"

    ' === 原紙ファイルの読み込み ===
    Set wbVer52 = Workbooks.Open(pathVer52 & "工数取得ver.5.2（大宮）yymm.xlsm")
    Set wsTarget = wbVer52.Worksheets("都度対応項目（時間）")

    wsTarget.Range("I6").Value = 年
    wsTarget.Range("I7").Value = 月

    fileNameSource = Dir(pathTemplate & "月次工数集計シード" & yyyymm & "*.xlsm")
    If fileNameSource = "" Then
        Call ErrorRollback("原紙（工数集計シード）ファイルが見つかりません。")
        Exit Sub
    End If

    Set wbSource = Workbooks.Open(pathTemplate & fileNameSource)
    Set wsSource = wbSource.Worksheets("工数取得-都度対応項目（時間）")

    ' === H10/H14/H18/H22 の数値転記 ===
    With wsTarget
        .Range("H10").NumberFormat = "[h]:mm"
        .Range("H10").Value = wsSource.Range("H10").Value
        .Range("H14").NumberFormat = "[h]:mm"
        .Range("H14").Value = wsSource.Range("H14").Value
        .Range("H18").NumberFormat = "[h]:mm"
        .Range("H18").Value = wsSource.Range("H18").Value
        .Range("H22").NumberFormat = "[h]:mm"
        .Range("H22").Value = wsSource.Range("H22").Value
    End With

    ' === O11:EH64 → O10:EH63 転記 ===
    wsSource.Range("O11:EH64").Copy
    wsTarget.Range("O10").PasteSpecial Paste:=xlPasteValues

    ' === ファイル名変更（内部用） ===
    yy = Right(年, 2)
    newFileName = "工数取得ver.5.2（大宮）" & yy & 月 & ".xlsm"
    wbVer52.Title = newFileName
    wbVer52.Windows(1).Caption = newFileName

    WriteLog "✅ 工数取得ver.5.2 転記完了：" & newFileName

    MsgBox "作業が完了しました。" & vbCrLf & _
           "このファイルは一時的に '" & newFileName & "' に名前を変更しました。" & vbCrLf & _
           "必ず「名前を付けて保存」してください。", vbInformation, "完了"

    ' 🔒 旧コード削除：フォルダ自動オープンなし
    wbSource.Close SaveChanges:=False

    Exit Sub

ERR_HANDLER:
    Call ErrorRollback("Process_Ver52でエラー：" & Err.Description)
End Sub


















