Sub 工数集計貼り付け()

    Dim userDate As String
    Dim yyyyMM As String
    Dim filePath As String
    Dim targetWb As Workbook
    Dim targetWs As Worksheet
    Dim categorySheets As Variant
    Dim categoryCodes As Variant
    Dim sourceWs As Worksheet
    Dim r As Long, lastRow As Long
    Dim dict As Object
    Dim key As String
    Dim dayVal As Long
    Dim shift As String
    Dim kousu As String
    Dim amount As Double
    Dim colOffset As Long
    Dim baseCol As Long
    Dim baseRow As Long
    Dim targetCol As Long
    Dim targetRow As Long
    Dim i As Long

    ' ===============================
    ' ユーザーに日付を入力させる
    ' ===============================
    userDate = InputBox("日報の日付を入力してください (フォーマット: YYYYMMDD)", "日付入力")
    If userDate = "" Or Not IsNumeric(userDate) Or Len(userDate) <> 8 Then
        MsgBox "有効な日付を入力してください（フォーマット: YYYYMMDD）", vbExclamation, "エラー"
        Exit Sub
    End If

    ' ↓ ここがポイント！処理する月を "実際の処理対象日" から取得
    yyyyMM = Left(userDate, 6) ' 入力日からYYYYMM取得
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
    ' 集計するシートとカテゴリ記号の定義
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

        ' -------------------------------
        ' 1. データ収集と合計
        ' -------------------------------
        For r = 2 To lastRow
            If Trim(sourceWs.Cells(r, 1).Value) <> "" And Trim(sourceWs.Cells(r, 2).Value) <> "" Then
                dayVal = CLng(sourceWs.Cells(r, 1).Value)
                shift = Trim(sourceWs.Cells(r, 2).Value)
                kousu = Trim(sourceWs.Cells(r, 4).Value)
                amount = Val(sourceWs.Cells(r, 5).Value) ' ← 合計済みの工数がE列にある！

                If IsNumeric(dayVal) And dayVal >= 1 And dayVal <= 31 Then
                    key = dayVal & "_" & shift & "_" & kousu
                    If dict.exists(key) Then
                        dict(key) = dict(key) + amount
                    Else
                        dict.Add key, amount
                    End If
                End If
            End If
        Next r

        ' -------------------------------
        ' 2. データを書き込み
        ' -------------------------------
        For Each key In dict.keys
            Dim parts() As String
            parts = Split(key, "_")
            dayVal = CLng(parts(0))
            shift = parts(1)
            kousu = parts(2)
            amount = dict(key)

            ' 列の決定
            baseCol = (dayVal - 1) * 10 + 6 ' F列=6 が1日目起点
            If shift = "日" Then
                colOffset = 0
            ElseIf shift = "夜" Then
                colOffset = 5
            Else
                GoTo SkipKey ' シフト不明はスキップ
            End If

            ' カテゴリごとの列（「客」「害」「他」「K」）
            ' 5行目（行番号=5）にそれぞれのカテゴリ文字がある列を探す
            targetCol = -1
            For j = 0 To 4 ' その日分の5列チェック
                If targetWs.Cells(5, baseCol + colOffset + j).Value = categoryCodes(i) Then
                    targetCol = baseCol + colOffset + j
                    Exit For
                End If
            Next j

            If targetCol = -1 Then
                GoTo SkipKey ' 対応する列が見つからない場合はスキップ
            End If

            ' 行の決定（工数番号が書かれているのはE7:E53）
            For targetRow = 7 To 53
                If Trim(targetWs.Cells(targetRow, 5).Value) = kousu Then Exit For
            Next targetRow

            ' 値の書き込み（上書きではなく累計して加算）
            If IsNumeric(targetWs.Cells(targetRow, targetCol).Value) Then
                targetWs.Cells(targetRow, targetCol).Value = targetWs.Cells(targetRow, targetCol).Value + amount
            Else
                targetWs.Cells(targetRow, targetCol).Value = amount
            End If

SkipKey:
        Next key

        Set dict = Nothing
    Next i

    MsgBox "工数の集計・転記が完了しました！", vbInformation, "完了"
    
    ' （必要であれば保存もここで追加可能）
    ' targetWb.Save
    ' targetWb.Close

End Sub
