Option Explicit

' ---start module 3
' fillBlanksAtFirstSheet の firstSheet を secondSheet に変え、安全規準に関するものを削除した
Sub fillBlanksAtSecondSheet()
    Dim offsetCol As Integer: offsetCol = 0 ' すべての遊具の最初のページまでの距離
    Dim offsetCol_firstPageOfYuugu As Integer: offsetCol_firstPageOfYuugu = 0 ' 遊具ごとの最初のページまでの距離。マイナス値

    Dim val_Yuugu As String: val_Yuugu = ""
    Dim val_Sougou As String
    Dim val_Rekka As Integer: val_Rekka = 1
    Dim val_Hazard As Integer: val_Hazard = 0
    Dim val_ZentaiComment As String: val_ZentaiComment = ""
    Dim val_Buzai As String
    Dim val_Zairyou As String

    Dim itr_rowOffset As Integer
    Dim itr_colOffset As Integer

    Sheets(sheetName_second).Activate

    Do While (Cells(row_place, firstCol_place + offsetCol) <> "")
        ' ---start  遊具の個別情報
        For itr_rowOffset = 0 To height_Kobetsu * 2 Step height_Kobetsu
            For itr_colOffset = 0 To width_Kobetsu Step width_Kobetsu

                ' ---start 部材名をセットまたは更新
                If (Cells(firstRow_secondSheet_Buzai + itr_rowOffset, firstCol_secondSheet_Buzai + itr_colOffset + offsetCol) <> "") Then
                    val_Buzai = Cells(firstRow_secondSheet_Buzai + itr_rowOffset, firstCol_secondSheet_Buzai + itr_colOffset + offsetCol)
                Else
                    ' 写真番号が書いてないなら、部材名などは転記しない
                    If (Cells(firstRow_secondSheet_img_Kobetsu + itr_rowOffset, firstCol_secondSheet_img_Kobetsu + itr_colOffset + offsetCol) <> "") Then
                        Cells(firstRow_secondSheet_Buzai + itr_rowOffset, firstCol_secondSheet_Buzai + itr_colOffset + offsetCol) = val_Buzai
                    End If
                End If
                ' ---end 部材名をセットまたは更新

                ' ---start 部材名ごとに、個別コメントを全体コメントに追記したり、最悪判断を更新したり
                Dim val_KobetsuComment As String
                val_KobetsuComment = Cells(firstRow_secondSheet_KobetsuComment + itr_rowOffset, firstCol_secondSheet_KobetsuComment + itr_colOffset + offsetCol)

                Dim val_Rekka_Kobetsu As Integer
                val_Rekka_Kobetsu = convertEvalStringToInt(Cells(firstRow_secondSheet_Rekka_Kobetsu + itr_rowOffset, firstCol_secondSheet_Rekka_Kobetsu + itr_colOffset + offsetCol))
                
                If (val_Rekka_Kobetsu > val_Rekka) Then
                    val_Rekka = val_Rekka_Kobetsu
                End If

                If (val_Rekka_Kobetsu = 1 And val_KobetsuComment = "") Then
                    Cells(firstRow_secondSheet_KobetsuComment + itr_rowOffset, firstCol_secondSheet_KobetsuComment + itr_colOffset + offsetCol) = "指摘なし。"
                End If

                If (val_Rekka_Kobetsu > 1 And InStr(val_ZentaiComment, val_KobetsuComment) = 0) Then
                    val_ZentaiComment = val_ZentaiComment + val_KobetsuComment
                    End If
                ' ---end 部材名ごとに、個別コメントを全体コメントに追記したり、最悪判断を更新したり

                ' ---start 材料をセットまたは更新
                If (Cells(firstRow_secondSheet_Zairyou + itr_rowOffset, firstCol_secondSheet_Zairyou + itr_colOffset + offsetCol) <> "") Then
                    val_Zairyou = Cells(firstRow_secondSheet_Zairyou + itr_rowOffset, firstCol_secondSheet_Zairyou + itr_colOffset + offsetCol)
                Else
                    If (Cells(firstRow_secondSheet_img_Kobetsu + itr_rowOffset, firstCol_secondSheet_img_Kobetsu + itr_colOffset + offsetCol) <> "") Then
                        Cells(firstRow_secondSheet_Zairyou + itr_rowOffset, firstCol_secondSheet_Zairyou + itr_colOffset + offsetCol) = val_Zairyou
                    End If
                End If
                ' ---end 材料をセットまたは更新

            Next itr_colOffset
        Next itr_rowOffset
        ' ---end 遊具の個別情報

        ' ---start ページ上部にある遊具の全体情報
        If (Cells(row_Yuugu, firstCol_Yuugu + offsetCol) <> "") Then ' 遊具名が空欄じゃない = 最初のページ
            val_Yuugu = Cells(row_Yuugu, firstCol_Yuugu + offsetCol)
            val_ZentaiComment = Cells(row_secondSheet_ZentaiComment, firstCol_secondSheet_ZentaiComment + offsetCol) + val_ZentaiComment
            If (Cells(row_place, firstCol_place + offsetCol + width_page) <> "") Then
                If (Cells(row_Yuugu, firstCol_Yuugu + offsetCol + width_page) <> "") Then
                    If (val_ZentaiComment = "") Then: val_ZentaiComment = "指摘なし。"
                    If (val_Rekka = 1) And (val_Hazard < 3) And (InStr(val_ZentaiComment, "指摘なし。") = 0) Then: val_ZentaiComment = val_ZentaiComment + "指摘なし。"
                    Call setAndResetLastEval(val_Sougou, val_Rekka, val_Hazard, val_ZentaiComment, offsetCol_firstPageOfYuugu, offsetCol)
                End If
            Else
                Call setAndResetLastEval(val_Sougou, val_Rekka, val_Hazard, val_ZentaiComment, offsetCol_firstPageOfYuugu, offsetCol)
            End If

        Else
            Cells(row_Yuugu, firstCol_Yuugu + offsetCol) = val_Yuugu + val_YuuguSuffix
            Cells(row_Suuryou, firstCol_Suuryou + offsetCol) = "" ' 数量は手書きで、初期値は1に設定しているので、つづきページの場合はそれを空欄に
            offsetCol_firstPageOfYuugu = offsetCol_firstPageOfYuugu - width_page

            If (Cells(row_Yuugu, firstCol_Yuugu + offsetCol + width_page) <> "") Then
                If (val_ZentaiComment = "") Then: val_ZentaiComment = "指摘なし。"
                Call setAndResetLastEval(val_Sougou, val_Rekka, val_Hazard, val_ZentaiComment, offsetCol_firstPageOfYuugu, offsetCol)
             Else
                Call setAndResetLastEval(val_Sougou, val_Rekka, val_Hazard, val_ZentaiComment, offsetCol_firstPageOfYuugu, offsetCol)
            End If

        End If
        ' ---end  ページ上部にある遊具の全体情報

        offsetCol = offsetCol + width_page
    Loop ' Do While
End Sub


Sub setAndResetLastEval(ByRef val_Sougou As String, ByRef val_Rekka As Integer, ByRef val_Hazard As Integer, ByRef val_ZentaiComment As String, ByRef offsetCol_firstPageOfYuugu As Integer, offsetCol)
    val_Sougou = decideSougou(val_Rekka, val_Hazard)
    Call setLastEval(val_Sougou, val_Rekka, val_Hazard, val_ZentaiComment, offsetCol, offsetCol_firstPageOfYuugu)
            
    ' 変数の中身をリセット
    val_Rekka = 1
    val_Hazard = 0
    val_ZentaiComment = ""
    offsetCol_firstPageOfYuugu = 0
End Sub


Function convertEvalStringToInt(val_Rekka)
    If (val_Rekka = "a") Then
        convertEvalStringToInt = 1
    ElseIf (val_Rekka = "b") Then
        convertEvalStringToInt = 2
    ElseIf (val_Rekka = "c") Then
        convertEvalStringToInt = 3
    ElseIf (val_Rekka = "d") Then
        convertEvalStringToInt = 4
    Else
        convertEvalStringToInt = 0
    End If
End Function


Function convertEvalIntToString(val_Rekka)
    If (val_Rekka = 1) Then
        convertEvalIntToString = "a"
    ElseIf (val_Rekka = 2) Then
        convertEvalIntToString = "b"
    ElseIf (val_Rekka = 3) Then
        convertEvalIntToString = "c"
    Else
        convertEvalIntToString = "d"
    End If
End Function


Function decideSougou(val_Rekka, val_Hazard)
    If (val_Rekka = 4) Then
        decideSougou = "D"
    ElseIf (val_Rekka = 3 Or val_Hazard = 3) Then
        decideSougou = "C"
    ElseIf (val_Rekka = 2 Or val_Hazard = 2 Or val_Hazard = 1) Then
        decideSougou = "B"
    Else
        decideSougou = "A"
    End If
End Function


Sub setLastEval(val_Sougou, val_Rekka, val_Hazard, val_ZentaiComment, offsetCol, offsetCol_firstPageOfYuugu)
    Cells(row_Sougou, firstCol_Sougou + offsetCol + offsetCol_firstPageOfYuugu) = val_Sougou
    Cells(row_secondSheet_ZentaiComment, firstCol_secondSheet_ZentaiComment + offsetCol + offsetCol_firstPageOfYuugu) = val_ZentaiComment

    If (val_Rekka = 1) Then: Cells(row_secondSheet_Rekka_Shochi, firstCol_secondSheet_Rekka_Shochi + offsetCol + offsetCol_firstPageOfYuugu) = "/"
End Sub
' ---end module 3
