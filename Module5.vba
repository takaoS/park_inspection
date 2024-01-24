Option Explicit

Sub fillNoAtAllSheet()
    Dim YuuguNo As Integer: YuuguNo = 0 ' 遊具(施設)の番号

    Call fillNo(sheetName_first, YuuguNo)
    Call fillNo(sheetName_second, YuuguNo)
End Sub

Sub fillNo(sheetName, ByRef YuuguNo As Integer)
    Dim offsetCol As Integer: offsetCol = 0 ' すべての遊具の最初のページまでの距離

    Sheets(sheetName).Activate
    Do While (Cells(row_place, firstCol_place + offsetCol) <> "")
        Dim val_Yuugu As String: val_Yuugu = Cells(row_Yuugu, firstCol_Yuugu + offsetCol)
        If (InStr(val_Yuugu, val_YuuguSuffix) = 0) And (val_Yuugu <> "") Then
            YuuguNo = YuuguNo + 1
            Cells(row_YuuguNo, firstCol_YuuguNo + offsetCol) = YuuguNo
        Else
            Cells(row_YuuguNo, firstCol_YuuguNo + offsetCol) = YuuguNo
        End If

        offsetCol = offsetCol + width_page
    Loop
End Sub
