Sub 工数集計貼り付け()

    Dim reportDate As String
    Dim yyyyMM As String
    Dim filePath As String
    Dim targetWb As Workbook
    Dim targetWs As Worksheet
    Dim categorySheets As Variant
    Dim categoryCodes As Variant
    Dim sourceWs As Worksheet
    Dim r As Long, lastRow As Long
    Dim dict As Object
    Dim dictKey As Variant
    Dim parts() As String
    Dim dayVal As Long
    Dim shift As String
    Dim kousu As String
    Dim amount As Double
    Dim colOffset As Long
    Dim baseCol As Long
    Dim baseRow As Long
    Dim targetCol As Long
    Dim targetRow As Long
    Dim i As Long, j As Long

    ' ===============================
    ' reportDate（命名変数）から日付取得
    ' ===============================
    On Error Resume Next
    reportDate = Evaluate(ThisWorkbook.Names("reportDate").RefersTo)
    On Error GoTo 0

    If reportDate = "" Then
        MsgBox "日付が取得できません。先にタイトル識別処理を実行してください。", vbExclamation, "エラー"
        Exit Sub
    End If

    yyyyMM = Left(reportDate, 6)
    filePath = "C:\Users\lis105\Desktop\06. 日報\④工数集計_日報_" & yyyyMM & ".xlsx"

    ' ===============================
    ' 対象ファイルを開く（存在確認）
    ' ===============================
    If Dir(filePath) = "" Then
        MsgBox "集計先のファイルが見つかりません：" & vbCrLf & filePath, vbCritical, "ファイル未発見"
        Exit Sub
    End If

    Set targetWb = Workbooks.Open(filePath)
    Set targetWs = targetWb.Sheets("日報")

    ' ===============================
    ' カテゴリと記号の定義
    ' ===============================
    categorySheets = Array("顧客対応", "障害対応", "その他", "KIX11業務")
    categoryCodes = Array("客", "害", "他", "K")

    ' ===============================
    ' 各カテゴリシートを処理
    ' ===============================
    For i = LBound(categorySheets) To UBound(categorySheets)
        Set sourceWs = ThisWorkbook.Sheets(categorySheets(i))
        Set dict = CreateObject("Scripting.Dictionary")
        lastRow = sourceWs.Cells(sourceWs.Rows.Count, "A").End(xlUp).Row

        ' データ収集と集計
        For r = 2 To lastRow
            If Trim(sourceWs.Cells(r, 1).Value) <> "" And Trim(sourceWs.Cells(r, 2).Value) <> "" Then
                dayVal = CLng(sourceWs.Cells(r, 1).Value)
                shift = Trim(sourceWs.Cells(r, 2).Value)
                kousu = Trim(sourceWs.Cells(r, 4).Value)
                amount = Val(sourceWs.Cells(r, 5).Value)

                If IsNumeric(dayVal) And dayVal >= 1 And dayVal <= 31 Then
                    dictKey = dayVal & "_" & shift & "_" & kousu
                    If dict.exists(dictKey) Then
                        dict(dictKey) = dict(dictKey) + amount
                    Else
                        dict.Add dictKey, amount
                    End If
                End If
            End If
        Next r

        ' 集計結果を貼り付け
        For Each dictKey In dict.Keys
            parts = Split(dictKey, "_")
            dayVal = CLng(parts(0))
            shift = parts(1)
            kousu = parts(2)
            amount = dict(dictKey)

            ' 列の決定
            baseCol = (dayVal - 1) * 10 + 6 ' F列起点（6列）
            If shift = "日" Then
                colOffset = 0
            ElseIf shift = "夜" Then
                colOffset = 5
            Else
                GoTo SkipKey
            End If

            ' 5行目のカテゴリ記号（客、害、他、K）から列を特定
            targetCol = -1
            For j = 0 To 4
                If targetWs.Cells(5, baseCol + colOffset + j).Value = categoryCodes(i) Then
                    targetCol = baseCol + colOffset + j
                    Exit For
                End If
            Next j

            If targetCol = -1 Then GoTo SkipKey

            ' 工数番号（E列）から行位置を特定
            For targetRow = 7 To 53
                If Trim(targetWs.Cells(targetRow, 5).Value) = kousu Then Exit For
            Next targetRow

            ' 累計で上書き（空白なら新規値）
            If IsNumeric(targetWs.Cells(targetRow, targetCol).Value) Then
                targetWs.Cells(targetRow, targetCol).Value = targetWs.Cells(targetRow, targetCol).Value + amount
            Else
                targetWs.Cells(targetRow, targetCol).Value = amount
            End If

SkipKey:
        Next dictKey

        Set dict = Nothing
    Next i

    MsgBox "工数の集計・転記が完了しました！", vbInformation, "完了"

    ' （必要に応じて保存・閉じる）
    ' targetWb.Save
    ' targetWb.Close

End Sub
