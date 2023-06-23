Attribute VB_Name = "mpb_vbascript_const"
Option Explicit

' シートを参照せず､スクリプト内でも変更されることのない定数を定義

' Googleドライブのパス
Public MPB_WORK_DIRECTORY_PATH As String

' 球団略称→球団名
Public DICT_TEAMNAME As New Dictionary

' 基礎スペ率
Public BASE_ACCIDENT_RATE As Long

' 球団略称→ヤ戦病院適用分
Public DICT_ACCIDENT_HDCP As New Dictionary

' スペ査定値→係数
Public DICT_ACCIDENT_COEFFICIENT As New Dictionary

' 表スペ抽選用辞書
Public DICT_ACCIDENT_LENGTH_RATE As New Dictionary

' 裏スペ抽選用二次元辞書
Public DICT_ACCIDENT_MARGIN_DICT As New Dictionary
Public DICT_ACCIDENT_MARGIN_S As New Dictionary
Public DICT_ACCIDENT_MARGIN_A As New Dictionary
Public DICT_ACCIDENT_MARGIN_B As New Dictionary
Public DICT_ACCIDENT_MARGIN_C As New Dictionary
Public DICT_ACCIDENT_MARGIN_D As New Dictionary
Public DICT_ACCIDENT_MARGIN_E As New Dictionary
Public DICT_ACCIDENT_MARGIN_F As New Dictionary
Public DICT_ACCIDENT_MARGIN_G As New Dictionary
Public DICT_ACCIDENT_MARGIN_n As New Dictionary

Public Function Definition()

    MPB_WORK_DIRECTORY_PATH = "C:\Users\TaiNo\マイドライブ\MPB\1-まる"
    
    With DICT_TEAMNAME
        .Add "G", "ジャイアンツ"
        .Add "M", "マリーンズ"
        .Add "T", "タイガース"
        .Add "L", "ライオンズ"
        .Add "E", "イーグルス"
    End With
    
    BASE_ACCIDENT_RATE = 0.01
    
    With DICT_ACCIDENT_HDCP
        .Add "G", 1#
        .Add "M", 1#
        .Add "T", 1#
        .Add "L", 1#
        .Add "E", 1#
    End With
    
    With DICT_ACCIDENT_COEFFICIENT
        .Add "S", 0.01
        .Add "A", 0.3
        .Add "B", 0.5
        .Add "C", 0.8
        .Add "D", 1#
        .Add "E", 1.2
        .Add "F", 2#
        .Add "G", 4#
        .Add "n", 0#
    End With
    
    With DICT_ACCIDENT_LENGTH_RATE
        .Add 1, 30#
        .Add 2, 49#
        .Add 5, 10.5
        .Add 8, 7
        .Add 24, 3.5
    End With
    
    With DICT_ACCIDENT_MARGIN_DICT
        .Add "S", DICT_ACCIDENT_MARGIN_S
        .Add "A", DICT_ACCIDENT_MARGIN_A
        .Add "B", DICT_ACCIDENT_MARGIN_B
        .Add "C", DICT_ACCIDENT_MARGIN_C
        .Add "D", DICT_ACCIDENT_MARGIN_D
        .Add "E", DICT_ACCIDENT_MARGIN_E
        .Add "F", DICT_ACCIDENT_MARGIN_F
        .Add "G", DICT_ACCIDENT_MARGIN_G
    End With
    
    With DICT_ACCIDENT_MARGIN_S
        .Add -1, 30#
        .Add 0, 70#
    End With
    
    With DICT_ACCIDENT_MARGIN_A
        .Add -1, 30#
        .Add 0, 40#
        .Add 1, 30#
    End With
    
    With DICT_ACCIDENT_MARGIN_B
        .Add -1, 30#
        .Add 0, 40#
        .Add 1, 20#
        .Add 2, 10#
    End With
    
    With DICT_ACCIDENT_MARGIN_C
        .Add -1, 30#
        .Add 0, 40#
        .Add 1, 15#
        .Add 2, 10#
        .Add 3, 5#
    End With
    
    With DICT_ACCIDENT_MARGIN_D
        .Add -1, 30#
        .Add 0, 40#
        .Add 1, 12#
        .Add 2, 9#
        .Add 3, 6#
        .Add 4, 3#
    End With
    
    With DICT_ACCIDENT_MARGIN_E
        .Add -1, 30#
        .Add 0, 40#
        .Add 1, 10#
        .Add 2, 8#
        .Add 3, 6#
        .Add 4, 4#
        .Add 5, 2#
    End With
    
    With DICT_ACCIDENT_MARGIN_F
        .Add -1, 30#
        .Add 0, 40#
        .Add 1, 8.57
        .Add 2, 7.14
        .Add 3, 5.71
        .Add 4, 4.29
        .Add 5, 2.86
        .Add 6, 1.43
    End With
    
    With DICT_ACCIDENT_MARGIN_G
        .Add -1, 30#
        .Add 0, 40#
        .Add 1, 7.5
        .Add 2, 6.43
        .Add 3, 5.36
        .Add 4, 4.29
        .Add 5, 3.21
        .Add 6, 2.14
        .Add 7, 1.07
    End With

End Function
