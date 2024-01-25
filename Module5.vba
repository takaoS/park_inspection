Option Explicit

Sub fillNumAtAllSheet()
    Dim YuuguNum As Integer: YuuguNum = 0 ' 遊具(施設)の番号

    Call fillNum(sheetName_first, YuuguNum)
    Call fillNum(sheetName_second, YuuguNum)
End Sub

Sub fillNum(sheetName, ByRef YuuguNum As Integer)
    Dim offsetCol As Integer: offsetCol = 0 ' すべての遊具の最初のページまでの距離

    Sheets(sheetName).Activate
    Do While (Cells(row_place, firstCol_place + offsetCol) <> "")
        Dim val_Yuugu As String: val_Yuugu = Cells(row_Yuugu, firstCol_Yuugu + offsetCol)
        If (InStr(val_Yuugu, val_YuuguSuffix) = 0) And (val_Yuugu <> "") Then
            YuuguNum = YuuguNum + 1
            Cells(row_YuuguNum, firstCol_YuuguNum + offsetCol) = YuuguNum
        Else
            Cells(row_YuuguNum, firstCol_YuuguNum + offsetCol) = YuuguNum
        End If

        offsetCol = offsetCol + width_page
    Loop
End Sub
