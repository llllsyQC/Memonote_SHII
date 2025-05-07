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

以前から自分用に整理していた作業フローやツールを使用していましたが、
その後に更新された最新バージョンが反映されておらず、結果的に今回の6Aのログ漏れにつながりました。
今後は自作ツールも定期的に見直し、最新版への更新を徹底します。
今後は自作フローやツールも定期的に見直し・更新確認を行うことを習慣化し、常に最新版に基づいて作業できるようにします。
・今回の件を踏まえて、「月初提出ファイルチェックシート」にtarget of event列による対象箇所のフィルター確認（No.16）を新設しました。
・また、使用ツール（merge_and_translate.html）が最新版（20250112）であることの明記を備考に追記し、チェックの徹底を図っています。