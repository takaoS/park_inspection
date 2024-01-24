Option Explicit

' ---start module4
Sub insertImagesAtAllSheet()
    Dim imgFolder As String
    imgFolder = InputBox("画像ファイルが保存してある場所を入力", "ユーザー入力", placeHolder_imgFolder)
    
    ' 表紙ページの写真挿入
    Call insertImage(sheetName_Top, row_topSheet_imgNo, row_topSheet_imgPlace, col_topSheet_imgNo, col_topSheet_imgPlace, offsetFromLeft_Top, offsetFromTop_Top, imgWidth_Top, imgHeight_Top, imgFolder)

    ' 最初の詳細ページの写真挿入
    Call insertInspectionImages(sheetName_first, row_firstSheet_ZentaiImgNo, row_firstSheet_ZentaiImgPlace, firstRow_firstSheet_KobetsuImgNo, firstRow_firstSheet_KobetsuImgPlace, _
        firstCol_firstSheet_ZentaiImgNo, firstCol_firstSheet_ZentaiImgPlace, firstCol_firstSheet_KobetsuImgNo, firstCol_firstSheet_KobetsuImgPlace, imgFolder)
        
    ' 2番目の詳細ページの写真挿入
    Call insertInspectionImages(sheetName_second, row_secondSheet_ZentaiImgNo, row_secondSheet_ZentaiImgPlace, firstRow_secondSheet_KobetsuImgNo, firstRow_secondSheet_KobetsuImgPlace, _
        firstCol_secondSheet_ZentaiImgNo, firstCol_secondSheet_ZentaiImgPlace, firstCol_secondSheet_KobetsuImgNo, firstCol_secondSheet_KobetsuImgPlace, imgFolder)
End Sub


Sub insertInspectionImages(sheetName, row_ZentaiImgNo, row_ZentaiImgPlace, firstRow_KobetsuImgNo, firstRow_KobetsuImgPlace, _
    firstCol_ZentaiImgNo, firstCol_ZentaiImgPlace, firstCol_KobetsuImgNo, firstCol_KobetsuImgPlace, imgFolder)
    Dim offsetCol As Integer: offsetCol = 0 ' すべての遊具の最初のページまでの距離

    Sheets(sheetName).Activate
    While (Cells(row_ZentaiImgNo, firstCol_ZentaiImgNo + offsetCol) <> "" Or Cells(firstRow_KobetsuImgNo, firstCol_KobetsuImgNo + offsetCol))
        Call insertImage(sheetName, row_ZentaiImgNo, row_ZentaiImgPlace, _
            firstCol_ZentaiImgNo + offsetCol, firstCol_ZentaiImgPlace + offsetCol, _
            offsetFromLeft_Zentai, offsetFromTop_Zentai, imgWidth_Zentai, imgHeight_Zentai, imgFolder)

        Dim itr_rowOffset As Integer
        Dim itr_colOffset As Integer
        For itr_rowOffset = 0 To height_Kobetsu * 2 Step height_Kobetsu
            For itr_colOffset = 0 To width_Kobetsu Step width_Kobetsu
                Call insertImage(sheetName, firstRow_KobetsuImgNo + itr_rowOffset, firstRow_KobetsuImgPlace + itr_rowOffset, _
                    firstCol_KobetsuImgNo + offsetCol + itr_colOffset, firstCol_KobetsuImgPlace + offsetCol + itr_colOffset, _
                    offsetFromLeft_Kobetsu, offsetFromTop_Kobetsu, imgWidth_Kobetsu, imgHeight_Kobetsu, imgFolder)
            Next itr_colOffset
        Next itr_rowOffset

        offsetCol = offsetCol + width_page
    Wend

End Sub

Sub insertImage(sheetName, row_ImgNo, row_imgPlace, col_ImgNo, col_imgPlace, offsetFromLeft, offsetFromTop, imgWidth, imgHeight, imgFolder)
    Dim imgNum As String: imgNum = Sheets(sheetName).Cells(row_ImgNo, col_ImgNo)

    If imgNum <> "" Then
        Dim imgPath As String: imgPath = imgFolder & "\" & imgPrefix & Format(imgNum, imgZeroPadding) & imgSuffix

        Sheets(sheetName).Activate

        Dim shapeObj As Object
        Set shapeObj = Sheets(sheetName).Shapes.AddPicture(Filename:=imgPath, LinkToFile:=False, SaveWithDocument:=True, Left:=Range(Cells(row_imgPlace, col_imgPlace).Address).Left + offsetFromLeft, Top:=Range(Cells(row_imgPlace, col_imgPlace).Address).Top + offsetFromTop, Width:=imgWidth, Height:=imgHeight)
    Else
        Sheets(sheetName).Cells(row_imgPlace, col_imgPlace) = "余白"
    End If
End Sub
' ---end module4

