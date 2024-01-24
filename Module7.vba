Option Explicit
Sub countFacilitiesOfAllYuuguAndShisetsu()
    Dim array_num_Yuugu_perEval(4) As Integer
    Dim array_num_Shisetsu_perEval(4) As Integer
    Dim i As Integer

    For i = 0 To 3
        array_num_Yuugu_perEval(i) = 0
        array_num_Shisetsu_perEval(i) = 0
    Next i

    Call countFacilities(startRow_sheetGaiyou_Yuugu, array_num_Yuugu_perEval())
    Call countFacilities(startRow_sheetGaiyou_Shisetsu, array_num_Shisetsu_perEval())

    For i = 0 To 3
        Sheets(sheetName_Gaiyou).Cells(startRow_sheetGaiyou_Syuukei + i, col_sheetGaiyou_Syuukei_Yuugu) = array_num_Yuugu_perEval(i)
        Sheets(sheetName_Gaiyou).Cells(startRow_sheetGaiyou_Syuukei + i, col_sheetGaiyou_Syuukei_Shisetsu) = array_num_Shisetsu_perEval(i)
    Next i
End Sub

Sub countFacilities(startRow, ByRef array_num_perEval() As Integer)
    Dim targetRow As Integer: targetRow = startRow

    Sheets(sheetName_Gaiyou).Activate
    Do While (Cells(targetRow, col_sheetGaiyou_Sougou) <> "")
        Dim val_Sougou As String: val_Sougou = Cells(targetRow, col_sheetGaiyou_Sougou)
        Dim val_Suuryou As Integer: val_Suuryou = Cells(targetRow, col_sheetGaiyou_Suuryou)

        If (val_Sougou = "A") Then
            array_num_perEval(0) = array_num_perEval(0) + val_Suuryou
        ElseIf (val_Sougou = "B") Then
            array_num_perEval(1) = array_num_perEval(1) + val_Suuryou
        ElseIf (val_Sougou = "C") Then
            array_num_perEval(2) = array_num_perEval(2) + val_Suuryou
        ElseIf (val_Sougou = "D") Then
            array_num_perEval(3) = array_num_perEval(3) + val_Suuryou
        Else
            ' NOTHING
        End If

        targetRow = targetRow + 1
    Loop

End Sub

