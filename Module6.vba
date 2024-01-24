Option Explicit

Sub postResults()
    Dim pageNum As Integer: pageNum = startPageNum

    Dim offsetCol As Integer: offsetCol = 0 ' すべての遊具の最初のページまでの距離
    Dim targetRow As Integer: targetRow = startRow_sheetGaiyou_Yuugu

    Sheets(sheetName_first).Activate
    Do While (Cells(row_place, firstCol_place + offsetCol) <> "")
        Dim val_ZentaiComment As String
        val_ZentaiComment = Cells(row_firstSheet_ZentaiComment, firstCol_firstSheet_ZentaiComment + offsetCol)
        
        If val_ZentaiComment <> "" Then
            Call postResultOfYuugu(targetRow, pageNum, offsetCol, val_ZentaiComment)
            targetRow = targetRow + 1
        End If
        
        pageNum = pageNum + 1
        offsetCol = offsetCol + width_page
    Loop

    offsetCol = 0
    targetRow = startRow_sheetGaiyou_Shisetsu

    Sheets(sheetName_second).Activate
    Do While (Cells(row_place, firstCol_place + offsetCol) <> "")
        val_ZentaiComment = Cells(row_secondSheet_ZentaiComment, firstCol_secondSheet_ZentaiComment + offsetCol)
        
        If val_ZentaiComment <> "" Then
            Call postResultOfShisetsu(targetRow, pageNum, offsetCol, val_ZentaiComment)
            targetRow = targetRow + 1
        End If
        
        pageNum = pageNum + 1
        offsetCol = offsetCol + width_page
    Loop
End Sub


Sub postResultOfYuugu(targetRow, pageNum, offsetCol, val_ZentaiComment)
    ' 施設No.のコピペ
    Sheets(sheetName_Gaiyou).Cells(targetRow, col_sheetGaiyou_ShisetsuNo) = Sheets(sheetName_first).Cells(row_YuuguNo, firstCol_YuuguNo + offsetCol)

    ' 遊具名のコピペ
    Sheets(sheetName_Gaiyou).Cells(targetRow, col_sheetGaiyou_Yuugu) = Sheets(sheetName_first).Cells(row_Yuugu, firstCol_Yuugu + offsetCol)

    ' 総合評価のコピペ
    Sheets(sheetName_Gaiyou).Cells(targetRow, col_sheetGaiyou_Sougou) = Sheets(sheetName_first).Cells(row_Sougou, firstCol_Sougou + offsetCol)

    ' 数量のコピペ
    Sheets(sheetName_Gaiyou).Cells(targetRow, col_sheetGaiyou_Suuryou) = Sheets(sheetName_first).Cells(row_Suuryou, firstCol_Suuryou + offsetCol)

    ' 劣化判定のコピペ
    Sheets(sheetName_Gaiyou).Cells(targetRow, col_sheetGaiyou_Rekka) = Sheets(sheetName_first).Cells(row_Rekka, firstCol_Rekka + offsetCol)
    
    ' ハザードレベルのコピペ
    Sheets(sheetName_Gaiyou).Cells(targetRow, col_sheetGaiyou_Hazard) = Sheets(sheetName_first).Cells(row_Hazard, firstCol_Hazard + offsetCol)

    ' 全体コメントのコピペ
    Sheets(sheetName_Gaiyou).Cells(targetRow, col_sheetGaiyou_ZentaiComment_first) = val_ZentaiComment

    ' ページ番号挿入
    Sheets(sheetName_Gaiyou).Cells(targetRow, col_sheetGaiyou_pageNum) = pageNum

    Sheets(sheetName_Gaiyou).Cells(targetRow + 1, col_sheetGaiyou_Yuugu) = "以下余白"
End Sub

Sub postResultOfShisetsu(targetRow, pageNum, offsetCol, val_ZentaiComment)
    ' 施設No.のコピペ
    Sheets(sheetName_Gaiyou).Cells(targetRow, col_sheetGaiyou_ShisetsuNo) = Sheets(sheetName_second).Cells(row_YuuguNo, firstCol_YuuguNo + offsetCol)

    ' 遊具名のコピペ
    Sheets(sheetName_Gaiyou).Cells(targetRow, col_sheetGaiyou_Yuugu) = Sheets(sheetName_second).Cells(row_Yuugu, firstCol_Yuugu + offsetCol)

    ' 総合評価のコピペ
    Sheets(sheetName_Gaiyou).Cells(targetRow, col_sheetGaiyou_Sougou) = Sheets(sheetName_second).Cells(row_Sougou, firstCol_Sougou + offsetCol)

    ' 数量のコピペ
    Sheets(sheetName_Gaiyou).Cells(targetRow, col_sheetGaiyou_Suuryou) = Sheets(sheetName_second).Cells(row_Suuryou, firstCol_Suuryou + offsetCol)

    ' 全体コメントのコピペ
    Sheets(sheetName_Gaiyou).Cells(targetRow, col_sheetGaiyou_ZentaiComment_second) = val_ZentaiComment

    ' ページ番号挿入
    Sheets(sheetName_Gaiyou).Cells(targetRow, col_sheetGaiyou_pageNum) = pageNum

    Sheets(sheetName_Gaiyou).Cells(targetRow + 1, col_sheetGaiyou_Yuugu) = "以下余白"
End Sub

