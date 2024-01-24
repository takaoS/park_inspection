Option Explicit

Public Const width_page = 11
Public Const height_Kobetsu = 11 ' 個別指摘の各枠の行数
Public Const width_Kobetsu = 5 ' 個別指摘の各枠の列数

Public Const row_place = 2
Public Const row_Yuugu = 3
Public Const row_YuuguNo = 3
Public Const row_Suuryou = 4
Public Const row_Sougou = 4

Public Const firstCol_place = 7
Public Const firstCol_Yuugu = 7
Public Const firstCol_YuuguNo = 4
Public Const firstCol_Suuryou = 7
Public Const firstCol_Sougou = 4

Public Const val_YuuguSuffix = "(つづき)"

' ---start 写真関係
Public Const imgPrefix = "DSCN"
Public Const imgSuffix = ".JPG"
Public Const imgZeroPadding = "0000"

Public Const placeHolder_imgFolder = "C:\Users\sekiguchi takao\OneDrive - 株式会社田中スポーツ設備\04 令和4年度（2022.4.1～2023.3.31）\＜点検業務＞\藤枝市 公園\写真\20230110"

Public Const imgWidth_Top = 310
Public Const imgHeight_Top = 233
Public Const offsetFromLeft_Top = 0
Public Const offsetFromTop_Top = 1

Public Const imgWidth_Zentai = 260
Public Const imgHeight_Zentai = 195
Public Const offsetFromLeft_Zentai = 42
Public Const offsetFromTop_Zentai = 1

Public Const imgWidth_Kobetsu = 167
Public Const imgHeight_Kobetsu = 125
Public Const offsetFromLeft_Kobetsu = 60
Public Const offsetFromTop_Kobetsu = 1
' ---end 写真関係

' ---start 表紙のシートについて
Public Const sheetName_Top = "表紙"
Public Const row_topSheet_imgNo = 11
Public Const col_topSheet_imgNo = 7
Public Const row_topSheet_imgPlace = 11
Public Const col_topSheet_imgPlace = 7
' ---end 表紙のシートについて

' ---start 概要シートについて
Public Const sheetName_Gaiyou = "概要"

Public Const startPageNum = 3
Public Const startRow_sheetGaiyou_Yuugu = 13
Public Const startRow_sheetGaiyou_Shisetsu = 26

Public Const startRow_sheetGaiyou_Syuukei = 5
Public Const col_sheetGaiyou_Syuukei_Yuugu = 14
Public Const col_sheetGaiyou_Syuukei_Shisetsu = 15

Public Const col_sheetGaiyou_ShisetsuNo = 2
Public Const col_sheetGaiyou_Yuugu = 3
Public Const col_sheetGaiyou_Suuryou = 5
Public Const col_sheetGaiyou_Sougou = 6
Public Const col_sheetGaiyou_Rekka = 7
Public Const col_sheetGaiyou_Hazard = 8
Public Const col_sheetGaiyou_pageNum = 16

Public Const col_sheetGaiyou_ZentaiComment_first = 9
Public Const col_sheetGaiyou_ZentaiComment_second = 7
' ---end 概要シートについて

' ---start 最初のシート(遊具)について
Public Const sheetName_first = "詳細(遊具)"

Public Const row_Rekka = 5
Public Const row_Hazard = 6
Public Const row_firstSheet_Rekka_Shochi = 5
Public Const row_Hazard_Shochi = 6

Public Const firstCol_Rekka = 4
Public Const firstCol_Hazard = 4
Public Const firstCol_firstSheet_Rekka_Shochi = 7
Public Const firstCol_Hazard_Shochi = 7

Public Const row_firstSheet_ZentaiComment = 9
Public Const firstCol_firstSheet_ZentaiComment = 8

Public Const row_firstSheet_ZentaiImgNo = 9
Public Const firstCol_firstSheet_ZentaiImgNo = 2
Public Const row_firstSheet_ZentaiImgPlace = 9
Public Const firstCol_firstSheet_ZentaiImgPlace = 2

Public Const firstRow_firstSheet_KobetsuImgNo = 22
Public Const firstCol_firstSheet_KobetsuImgNo = 6
Public Const firstRow_firstSheet_KobetsuImgPlace = 24
Public Const firstCol_firstSheet_KobetsuImgPlace = 2

Public Const firstRow_firstSheet_Buzai = 22
Public Const firstCol_firstSheet_Buzai = 3
Public Const firstRow_firstSheet_Zairyou = 23
Public Const firstCol_firstSheet_Zairyou = 3
Public Const firstRow_firstSheet_img_Kobetsu = 22
Public Const firstCol_firstSheet_img_Kobetsu = 6
Public Const firstRow_firstSheet_Rekka_Kobetsu = 23 ' ハザードレベルも同じセルに記入するので、ハザード判定と併用
Public Const firstCol_firstSheet_Rekka_Kobetsu = 6
Public Const firstRow_firstSheet_KobetsuComment = 31
Public Const firstCol_firstSheet_KobetsuComment = 3
' ---end 最初のシート(遊具)について

' ---start 2番目のシート(施設)について
Public Const sheetName_second = "詳細(施設)"

Public Const row_secondSheet_Rekka_Shochi = 4
Public Const firstCol_secondSheet_Rekka_Shochi = 10

Public Const row_secondSheet_ZentaiComment = 7
Public Const firstCol_secondSheet_ZentaiComment = 8

Public Const row_secondSheet_ZentaiImgNo = 7
Public Const firstCol_secondSheet_ZentaiImgNo = 2
Public Const row_secondSheet_ZentaiImgPlace = 7
Public Const firstCol_secondSheet_ZentaiImgPlace = 2

Public Const firstRow_secondSheet_KobetsuImgNo = 20
Public Const firstCol_secondSheet_KobetsuImgNo = 6
Public Const firstRow_secondSheet_KobetsuImgPlace = 22
Public Const firstCol_secondSheet_KobetsuImgPlace = 2

Public Const firstRow_secondSheet_Buzai = 20
Public Const firstCol_secondSheet_Buzai = 3
Public Const firstRow_secondSheet_Zairyou = 21
Public Const firstCol_secondSheet_Zairyou = 3
Public Const firstRow_secondSheet_img_Kobetsu = 20
Public Const firstCol_secondSheet_img_Kobetsu = 6
Public Const firstRow_secondSheet_Rekka_Kobetsu = 21
Public Const firstCol_secondSheet_Rekka_Kobetsu = 6
Public Const firstRow_secondSheet_KobetsuComment = 29
Public Const firstCol_secondSheet_KobetsuComment = 3
' ---end 2番目のシート(施設)について


Sub コメント補完と画像挿入()

    Call Module2.fillBlanksAtFirstSheet
    Call Module3.fillBlanksAtSecondSheet
    Call Module4.insertImagesAtAllSheet

End Sub

Sub 概要ページに転記して集計()

    Call Module5.fillNoAtAllSheet
    Call Module6.postResults
    Call Module7.countFacilitiesOfAllYuuguAndShisetsu

End Sub
