Attribute VB_Name = "global_values"
Option Explicit

' シートを参照せず､スクリプト内でも変更されることのない定数を定義

' Googleドライブのパス
Public MPB_WORK_DIRECTORY_PATH As String

' ローカル動確時のパス
Public LOCAL_WORK_DIRECTORY_PATH As String

' 球団略称→球団名
Public DICT_TEAM_NAME As New Dictionary

' 基礎スペ率
Public BASE_ACCIDENT_RATE As Single

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

' 投手スペ内容抽選用二次元辞書
' MPBニュース出力を ◇チーム名◇選手名選手XXX とする場合の XXX を定義
Public DICT_ACCIDENT_INFORMATION_PITCHER_DICT As New Dictionary
Public DICT_ACCIDENT_INFORMATION_PITCHER_1 As New Dictionary
Public DICT_ACCIDENT_INFORMATION_PITCHER_2 As New Dictionary
Public DICT_ACCIDENT_INFORMATION_PITCHER_5 As New Dictionary
Public DICT_ACCIDENT_INFORMATION_PITCHER_8 As New Dictionary
Public DICT_ACCIDENT_INFORMATION_PITCHER_24 As New Dictionary

' 野手スペ内容抽選用二次元辞書
' MPBニュース出力を ◇チーム名◇選手名選手XXX とする場合の XXX を定義
Public DICT_ACCIDENT_INFORMATION_FIELDER_DICT As New Dictionary
Public DICT_ACCIDENT_INFORMATION_FIELDER_1 As New Dictionary
Public DICT_ACCIDENT_INFORMATION_FIELDER_2 As New Dictionary
Public DICT_ACCIDENT_INFORMATION_FIELDER_5 As New Dictionary
Public DICT_ACCIDENT_INFORMATION_FIELDER_8 As New Dictionary
Public DICT_ACCIDENT_INFORMATION_FIELDER_24 As New Dictionary

Public Function Definition()

    MPB_WORK_DIRECTORY_PATH = "C:\Users\TaiNo\マイドライブ\MPB\1-まる"
    LOCAL_WORK_DIRECTORY_PATH = "C:\Users\TaiNo\Desktop\MPB\1-まる"

    With DICT_TEAM_NAME
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
        .Add 1, 36#
        .Add 2, 40#
        .Add 5, 12#
        .Add 8, 8#
        .Add 24, 4#
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

    With DICT_ACCIDENT_INFORMATION_PITCHER_DICT
        .Add 1, DICT_ACCIDENT_INFORMATION_PITCHER_1
        .Add 2, DICT_ACCIDENT_INFORMATION_PITCHER_2
        .Add 5, DICT_ACCIDENT_INFORMATION_PITCHER_5
        .Add 8, DICT_ACCIDENT_INFORMATION_PITCHER_8
        .Add 24, DICT_ACCIDENT_INFORMATION_PITCHER_24
    End With

    With DICT_ACCIDENT_INFORMATION_PITCHER_1
        .Add "選手が、再調整のため、次節は一度ベンチから外れるとのことです。", 1
        .Add "選手は、上肢のコンディション不良により、念のため次節のベンチから外れるとのことです。", 1
        .Add "選手は、下肢のコンディション不良により、念のため次節のベンチから外れるとのことです。", 1
        .Add "選手は、練習中に肘の違和感を訴えたため、念のため次節のベンチから外れるとのことです。", 1
        .Add "選手は、練習中に肩の違和感を訴えたため、念のため次節のベンチから外れるとのことです。", 1
        .Add "選手は、練習中に腰の違和感を訴えたため、念のため次節のベンチから外れるとのことです。", 1
        .Add "選手に、ピッチャーライナーを受けるアクシデント。念のため次節はベンチから外れるとのことです。", 1
        .Add "選手は、指にできたマメの影響で、念のため次節のベンチから外れるとのことです。", 1
        .Add "選手は、体調不良のため、コロナ特例で登録抹消されました。", 1
    End With

    With DICT_ACCIDENT_INFORMATION_PITCHER_2
        .Add "選手が上肢のコンディション不良とのこと。抹消は行わず、様子を見る方針です。", 1
        .Add "選手が下肢のコンディション不良とのこと。抹消は行わず、様子を見る方針です。", 1
        .Add "選手が肘の違和感を訴えたとのこと。抹消は行わず、様子を見る方針です。", 1
        .Add "選手が肩の違和感を訴えたとのこと。抹消は行わず、様子を見る方針です。", 1
        .Add "選手が腰の違和感を訴えたとのこと。抹消は行わず、様子を見る方針です。", 1
        .Add "選手の指にマメができたとのこと。抹消は行わず、様子を見る方針です。", 1
    End With

    With DICT_ACCIDENT_INFORMATION_PITCHER_5
        .Add "選手が上肢のコンディション不良とのこと。一度抹消し、様子を見る方針です。", 1
        .Add "選手が下肢のコンディション不良とのこと。一度抹消し、様子を見る方針です。", 1
        .Add "選手が肘の違和感を訴えたとのこと。一度抹消し、様子を見る方針です。", 1
        .Add "選手が肩の違和感を訴えたとのこと。一度抹消し、様子を見る方針です。", 1
        .Add "選手が腰の違和感を訴えたとのこと。一度抹消し、様子を見る方針です。", 1
    End With

    With DICT_ACCIDENT_INFORMATION_PITCHER_8
        .Add "選手は、上肢のコンディション不良のため、登録抹消し、治療に専念するとのことです。", 1
        .Add "選手は、下肢のコンディション不良のため、登録抹消し、治療に専念するとのことです。", 1
        .Add "選手は、肘の違和感を訴えたため、登録抹消し、治療に専念するとのことです。", 1
        .Add "選手は、肩の違和感を訴えたため、登録抹消し、治療に専念するとのことです。", 1
        .Add "選手は、腰痛のため、登録抹消し、治療に専念するとのことです。", 1
    End With

    With DICT_ACCIDENT_INFORMATION_PITCHER_24
        .Add "選手が、直近の登板の際に肘を痛め、緊急降板。近日中に手術を行うとのことで、今シーズン中の復帰は絶望的とみられます。", 1
        .Add "選手が、直近の登板の際に肩を痛め、緊急降板。近日中に手術を行うとのことで、今シーズン中の復帰は絶望的とみられます。", 1
    End With

    With DICT_ACCIDENT_INFORMATION_FIELDER_DICT
        .Add 1, DICT_ACCIDENT_INFORMATION_FIELDER_1
        .Add 2, DICT_ACCIDENT_INFORMATION_FIELDER_2
        .Add 5, DICT_ACCIDENT_INFORMATION_FIELDER_5
        .Add 8, DICT_ACCIDENT_INFORMATION_FIELDER_8
        .Add 24, DICT_ACCIDENT_INFORMATION_FIELDER_24
    End With

    With DICT_ACCIDENT_INFORMATION_FIELDER_1
        .Add "選手が、再調整のため、次節は一度ベンチから外れるとのことです。", 1
        .Add "選手は、上肢のコンディション不良により、念のため次節のベンチから外れるとのことです。", 1
        .Add "選手は、下肢のコンディション不良により、念のため次節のベンチから外れるとのことです。", 1
        .Add "選手は、練習中に太ももの違和感を訴えたため、念のため次節のベンチから外れるとのことです。", 1
        .Add "選手は、練習中に腰の違和感を訴えたため、念のため次節のベンチから外れるとのことです。", 1
        .Add "選手は、死球による打撲の影響で、念のため次節のベンチから外れるとのことです。", 1
        .Add "選手は、走塁中のアクシデントの影響で、念のため次節のベンチから外れるとのことです。", 1
        .Add "選手は、体調不良を訴えたため、コロナ特例で登録抹消されました。", 1
    End With

    With DICT_ACCIDENT_INFORMATION_FIELDER_2
        .Add "選手が上肢のコンディション不良とのこと。抹消は行わず、様子を見る方針です。", 1
        .Add "選手が下肢のコンディション不良とのこと。抹消は行わず、様子を見る方針です。", 1
        .Add "選手が太ももの違和感を訴えたとのこと。抹消は行わず、様子を見る方針です。", 1
        .Add "選手が腰の違和感を訴えたとのこと。抹消は行わず、様子を見る方針です。", 1
        .Add "選手は、死球を受け市内の病院を受診。抹消は行わず、様子を見る方針です。", 1
        .Add "選手が、走塁中のアクシデントで途中交代しました。抹消は行わず、様子を見る方針です。", 1
    End With

    With DICT_ACCIDENT_INFORMATION_FIELDER_5
        .Add "選手が上肢のコンディション不良とのこと。一度抹消し、様子を見る方針です。", 1
        .Add "選手が下肢のコンディション不良とのこと。一度抹消し、様子を見る方針です。", 1
        .Add "選手が太ももの違和感を訴えたとのこと。一度抹消し、様子を見る方針です。", 1
        .Add "選手が腰の違和感を訴えたとのこと。一度抹消し、様子を見る方針です。", 1
        .Add "選手は、死球を受け市内の病院を受診。一度抹消し、様子を見る方針です。", 1
        .Add "選手が、走塁中のアクシデントで途中交代しました。一度抹消し、様子を見る方針です。", 1
    End With

    With DICT_ACCIDENT_INFORMATION_FIELDER_8
        .Add "選手は、上肢のコンディション不良のため、登録抹消し、治療に専念するとのことです。", 1
        .Add "選手は、下肢のコンディション不良のため、登録抹消し、治療に専念するとのことです。", 1
        .Add "選手は、腰痛のため、登録抹消し、治療に専念するとのことです。", 1
        .Add "選手は、死球を受け骨折。登録抹消し、治療に専念するとのことです。", 1
        .Add "選手は、走塁中のアクシデントで途中交代、肉離れと診断。登録抹消し、治療に専念するとのことです。", 1
    End With

    With DICT_ACCIDENT_INFORMATION_FIELDER_24
        .Add "選手が、守備の際に膝を痛め病院に直行。前十字靭帯損傷と診断されました。今シーズン中の復帰は絶望的とみられます。", 1
        .Add "選手は、腰痛を訴え病院を受診したところ、椎間板ヘルニアと診断されました。近日中に手術を行うとのことで、今シーズン中の復帰は絶望的とみられます。", 1
    End With

End Function
