Sub 工数集計貼り付け_テスト版()

    Dim reportDate As String
    Dim userInput As String
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
    Dim targetCol As Long
    Dim targetRow As Long
    Dim i As Long, j As Long
    Dim clearDayColStart As Long

    ' ✅ 1. 日付入力（独立用）
    userInput = InputBox("日報の日付を入力してください（フォーマット: YYYYMMDD）", "日付入力")
    If userInput = "" Or Not IsNumeric(userInput) Or Len(userInput) <> 8 Then
        MsgBox "有効な日付を入力してください（例：20241229）", vbExclamation
        Exit Sub
    End If

    reportDate = userInput
    yyyyMM = Left(reportDate, 6)
    filePath = "C:\Users\lis105\Desktop\06. 日報\④工数集計_日報_" & yyyyMM & ".xlsx"

    ' ✅ 2. ファイルチェック
    If Dir(filePath) = "" Then
        MsgBox "集計先ファイルが見つかりません：" & vbCrLf & filePath, vbCritical
        Exit Sub
    End If

    Set targetWb = Workbooks.Open(filePath)
    Set targetWs = targetWb.Sheets("日報")

    ' ✅ 3. カテゴリ定義
    categorySheets = Array("顧客対応", "障害対応", "その他", "KIX11業務")
    categoryCodes = Array("客", "害", "他", "K")

    ' ✅ 4. 各カテゴリ処理
    For i = LBound(categorySheets) To UBound(categorySheets)
        Set sourceWs = ThisWorkbook.Sheets(categorySheets(i))
        Set dict = CreateObject("Scripting.Dictionary")
        lastRow = sourceWs.Cells(sourceWs.Rows.Count, "A").End(xlUp).Row

        ' ✅ 先に該当日・カテゴリ列をクリア（F〜O列基準で10列分）
        clearDayColStart = (CLng(reportDate) Mod 100 - 1) * 10 + 6

        For j = 0 To 9
            For targetRow = 7 To 53
                If targetWs.Cells(5, clearDayColStart + j).Value = categoryCodes(i) Then
                    targetWs.Cells(targetRow, clearDayColStart + j).Value = ""
                End If
            Next targetRow
        Next j

        ' ✅ 5. データ読み込みと集計
        For r = 2 To lastRow
            If Trim(sourceWs.Cells(r, 1).Value) <> "" And Trim(sourceWs.Cells(r, 2).Value) <> "" Then
                If IsNumeric(sourceWs.Cells(r, 1).Value) Then
                    dayVal = CLng(sourceWs.Cells(r, 1).Value)
                Else
                    GoTo SkipRow
                End If

                shift = Trim(sourceWs.Cells(r, 2).Value)
                kousu = Trim(sourceWs.Cells(r, 4).Value)
                amount = Val(sourceWs.Cells(r, 5).Value)

                If dayVal >= 1 And dayVal <= 31 Then
                    dictKey = dayVal & "_" & shift & "_" & kousu
                    If dict.exists(dictKey) Then
                        dict(dictKey) = dict(dictKey) + amount
                    Else
                        dict.Add dictKey, amount
                    End If
                End If
            End If
SkipRow:
        Next r

        ' ✅ 6. 結果を貼り付け
        For Each dictKey In dict.Keys
            parts = Split(dictKey, "_")
            If UBound(parts) < 2 Then GoTo SkipKey
            If Not IsNumeric(parts(0)) Then GoTo SkipKey

            dayVal = CLng(parts(0))
            shift = parts(1)
            kousu = parts(2)
            amount = dict(dictKey)

            baseCol = (dayVal - 1) * 10 + 6
            If shift = "日" Then
                colOffset = 0
            ElseIf shift = "夜" Then
                colOffset = 5
            Else
                GoTo SkipKey
            End If

            targetCol = -1
            For j = 0 To 4
                If targetWs.Cells(5, baseCol + colOffset + j).Value = categoryCodes(i) Then
                    targetCol = baseCol + colOffset + j
                    Exit For
                End If
            Next j

            If targetCol = -1 Then GoTo SkipKey

            For targetRow = 7 To 53
                If Trim(targetWs.Cells(targetRow, 5).Value) = kousu Then Exit For
            Next targetRow

            targetWs.Cells(targetRow, targetCol).Value = amount
SkipKey:
        Next dictKey

        Set dict = Nothing
    Next i

    MsgBox "✅ 工数の集計・転記が完了しました！", vbInformation

    ' （必要に応じて保存・閉じる）
    ' targetWb.Save
    ' targetWb.Close

End Sub
