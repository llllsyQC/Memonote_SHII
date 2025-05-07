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
    Dim targetCol As Long
    Dim targetRow As Long
    Dim i As Long, j As Long
    Dim clearDayColStart As Long

    ' ✅ reportDate（名称変数）取得
    On Error Resume Next
    reportDate = Evaluate(ThisWorkbook.Names("reportDate").RefersTo)
    On Error GoTo 0

    If reportDate = "" Then
        MsgBox "日付が取得できません。先にタイトル識別処理を実行してください。", vbExclamation
        Exit Sub
    End If

    yyyyMM = Left(reportDate, 6)
    filePath = "C:\Users\lis105\Desktop\06. 日報\④工数集計_日報_" & yyyyMM & ".xlsx"

    If Dir(filePath) = "" Then
        MsgBox "集計先ファイルが見つかりません：" & vbCrLf & filePath, vbCritical
        Exit Sub
    End If

    Set targetWb = Workbooks.Open(filePath)
    Set targetWs = targetWb.Sheets("日報")

    categorySheets = Array("顧客対応", "障害対応", "その他", "KIX11業務")
    categoryCodes = Array("客", "害", "他", "K")

    For i = LBound(categorySheets) To UBound(categorySheets)
        Set sourceWs = ThisWorkbook.Sheets(categorySheets(i))
        Set dict = CreateObject("Scripting.Dictionary")
        lastRow = sourceWs.Cells(sourceWs.Rows.Count, "A").End(xlUp).Row

        ' ✅ 対象日データの事前クリア
        clearDayColStart = (CLng(reportDate) Mod 100 - 1) * 10 + 6

        For j = 0 To 9
            For targetRow = 7 To 53
                If targetWs.Cells(5, clearDayColStart + j).Value = categoryCodes(i) Then
                    targetWs.Cells(targetRow, clearDayColStart + j).Value = ""
                End If
            Next targetRow
        Next j

        ' ✅ データ集計
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

        ' ✅ 集計結果を書き込み
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
    ' targetWb.Save
    ' targetWb.Close

End Sub
Sub 一括日報処理()
    Call 更新と精度向上したタイトル識別
    Call 分類データ貼り付け
    Call 工数集計貼り付け ' ✅ 新增这一行
End Sub

' ===== 粘贴 H10/H14/H18/H22 の工時数値 =====
With wsTarget
    .Range("H10").NumberFormat = "[h]:mm"
    .Range("H10").Value = ThisWorkbook.Sheets("工数取得-都度対応項目(時間)").Range("H10").Value

    .Range("H14").NumberFormat = "[h]:mm"
    .Range("H14").Value = ThisWorkbook.Sheets("工数取得-都度対応項目(時間)").Range("H14").Value

    .Range("H18").NumberFormat = "[h]:mm"
    .Range("H18").Value = ThisWorkbook.Sheets("工数取得-都度対応項目(時間)").Range("H18").Value

    .Range("H22").NumberFormat = "[h]:mm"
    .Range("H22").Value = ThisWorkbook.Sheets("工数取得-都度対応項目(時間)").Range("H22").Value
End With