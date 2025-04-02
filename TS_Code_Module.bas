Sub TS自動計算_完全版()

    Dim wsInput As Worksheet
    Dim wsOutput As Worksheet
    Dim lastRow As Long, i As Long
    Dim 品名 As String, 備考 As String
    Dim 数量 As Long
    Dim TS合計 As Double
    Dim 鏡面補正 As Double: 鏡面補正 = 1

    ' 初期設定
    Set wsInput = Worksheets("品目票")
    On Error Resume Next
    Set wsOutput = Worksheets("TS出力")
    If wsOutput Is Nothing Then
        Set wsOutput = Worksheets.Add
        wsOutput.Name = "TS出力"
    Else
        wsOutput.Cells.ClearContents
    End If
    On Error GoTo 0

    ' ヘッダー
    wsOutput.Range("A1:E1").Value = Array("品名", "数量", "カテゴリ", "TS時間（h）", "備考")

    lastRow = wsInput.Cells(wsInput.Rows.Count, "B").End(xlUp).Row
    TS合計 = 0
    Dim outRow As Long: outRow = 2

    For i = 2 To lastRow
        品名 = wsInput.Cells(i, 2).Value
        数量 = wsInput.Cells(i, 3).Value
        備考 = wsInput.Cells(i, 4).Value

        Dim TS As Double: TS = 0
        Dim カテゴリ As String: カテゴリ = ""

        ' カテゴリ判定と時間加算
        If InStr(品名, "E-PIN") > 0 Then
            TS = 数量 * 0.2
            カテゴリ = "エジェクタピン"
        ElseIf InStr(品名, "スライド") > 0 Then
            TS = 数量 * 3
            カテゴリ = "スライド"
        ElseIf InStr(品名, "センターピン") > 0 Then
            TS = 数量 * 0.5
            カテゴリ = "センターピン"
        ElseIf InStr(品名, "リターンピン") > 0 Then
            TS = 数量 * 1
            カテゴリ = "リターンピン"
        ElseIf InStr(品名, "食い切り") > 0 Or InStr(品名, "くいきり") > 0 Then
            TS = 数量 * 2
            カテゴリ = "食い切り"
        ElseIf InStr(品名, "ガイドピン") > 0 Then
            TS = 数量 * 0.5
            カテゴリ = "ガイドピン"
        ElseIf InStr(品名, "ガイドブッシュ") > 0 Then
            TS = 数量 * 0.5
            カテゴリ = "ガイドブッシュ"
        ElseIf InStr(品名, "スプリング") > 0 Or InStr(品名, "MSWT") > 0 Then
            TS = 数量 * 0.3
            カテゴリ = "スプリング"
        End If

        ' 鏡面補正フラグ
        If 鏡面補正 = 1 And InStr(備考, "鏡面") > 0 Then
            鏡面補正 = 1.3
        End If

        ' 出力
        If TS > 0 Then
            wsOutput.Cells(outRow, 1).Value = 品名
            wsOutput.Cells(outRow, 2).Value = 数量
            wsOutput.Cells(outRow, 3).Value = カテゴリ
            wsOutput.Cells(outRow, 4).Value = TS
            wsOutput.Cells(outRow, 5).Value = 備考
            outRow = outRow + 1
            TS合計 = TS合計 + TS
        End If
    Next i

    ' 合計と補正
    wsOutput.Cells(outRow + 1, 3).Value = "合計TS（補正前）"
    wsOutput.Cells(outRow + 1, 4).Value = TS合計
    wsOutput.Cells(outRow + 2, 3).Value = "鏡面補正係数"
    wsOutput.Cells(outRow + 2, 4).Value = 鏡面補正
    wsOutput.Cells(outRow + 3, 3).Value = "最終TS時間"
    wsOutput.Cells(outRow + 3, 4).Value = TS合計 * 鏡面補正

    MsgBox "TS自動計算が完了しました！", vbInformation

End Sub
