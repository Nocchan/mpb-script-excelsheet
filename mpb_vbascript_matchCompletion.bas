Attribute VB_Name = "mpb_vbascript_matchCompletion"
Dim season As String
Dim game As Integer
Dim section As Integer

Dim MPB_WORK_DIRECTORY_PATH As String
Dim DICT_TEAMNAME As New Dictionary

Dim BASE_ACCIDENT_RATE As Long
Dim DICT_ACCIDENT_HDCP As New Dictionary
Dim DICT_ACCIDENT_COEFFICIENT As New Dictionary

Dim DICT_ACCIDENT_LENGTH_RATE As New Dictionary
Dim DICT_ACCIDENT_MARGIN_DICT As New Dictionary
Dim DICT_ACCIDENT_MARGIN_S As New Dictionary
Dim DICT_ACCIDENT_MARGIN_A As New Dictionary
Dim DICT_ACCIDENT_MARGIN_B As New Dictionary
Dim DICT_ACCIDENT_MARGIN_C As New Dictionary
Dim DICT_ACCIDENT_MARGIN_D As New Dictionary
Dim DICT_ACCIDENT_MARGIN_E As New Dictionary
Dim DICT_ACCIDENT_MARGIN_F As New Dictionary
Dim DICT_ACCIDENT_MARGIN_G As New Dictionary
Dim DICT_ACCIDENT_MARGIN_n As New Dictionary

Sub matchCompletion()
    
    ' デバッグモード
    ' Call DebugMode

    ' 呼出元確認
    If Not IsScheduleSheet() Then
        Call MessageError("呼出元確認エラー", "matchCompletion")
        End
    End If
    
    Call Initialize
    
    If IsSectionCompleted() Then
        Call MakeMPBNewsSeasonEvent
        Call MakeMPBNewsOfThisSection
        Call MakeMPBNewsOfAccident
        Call MakeMPBNewsOfNextGame
    End If
    
    Call SavePictureOfSchedule
    Call SavePictureOfRanking
    
    Call ExitProcess

End Sub

' 定数・シート状態の初期化
Function Initialize()

    If Not debugModeFlg Then
        Application.ScreenUpdating = False
    End If
    
    Application.Calculate
    
    season = ActiveSheet.Cells(1, "A").Value
    game = WorksheetFunction.CountIf(ActiveSheet.Range("BA2:BA241"), 0) / 4
    section = WorksheetFunction.CountIf(ActiveSheet.Range("BA2:BA241"), 0) / 8
    
    Sheets(season & "_投手データ").Unprotect
    Sheets(season & "_野手データ").Unprotect
    
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
        .Add "n", DICT_ACCIDENT_MARGIN_n
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

' 終了時処理
Function ExitProcess()

    Sheets(season & "_投手データ").Protect AllowFormattingColumns:=True, AllowFormattingRows:=True
    Sheets(season & "_野手データ").Protect AllowFormattingColumns:=True, AllowFormattingRows:=True

    If Not debugModeFlg Then
        Application.ScreenUpdating = True
    End If
    
    End
    
End Function

' 節が完了してスペ判定を行える状態かを判定
Function IsSectionCompleted() As Boolean

    ' 消化試合数から節が完了していないことがわかるパターン
    If game <> section * 2 Then
        IsSectionCompleted = False
        Exit Function
    End If
    
    ' 節は完了しているが不正入力があるパターン
    If ActiveSheet.Cells(section * 8 + 3, "D").Value <> "" Or ActiveSheet.Cells(section * 8 + 7, "D").Value <> "" Or _
       ActiveSheet.Cells(section * 8 + 3, "F").Value <> "" Or ActiveSheet.Cells(section * 8 + 7, "F").Value <> "" Or _
       ActiveSheet.Cells(section * 8 + 3, "H").Value <> "" Or ActiveSheet.Cells(section * 8 + 7, "H").Value <> "" Then
        Call MessageError("不正入力エラー", "IsSectionCompleted")
        Call ExitProcess
    End If
    
    ' 開幕前または最終節後で予告先発を考える必要がないパターン
    If section = 0 Or section = 30 Then
        IsSectionCompleted = True
        Exit Function
    End If
    
    ' 予告先発が出揃っていないパターン
    If ActiveSheet.Cells(section * 8 + 2, "D").Value = "" Or ActiveSheet.Cells(section * 8 + 6, "D").Value = "" Or _
       ActiveSheet.Cells(section * 8 + 2, "H").Value = "" Or ActiveSheet.Cells(section * 8 + 6, "H").Value = "" Then
        Call MessageError("予告先発未完了エラー", "IsSectionCompleted")
        Call ExitProcess
    End If
    
    IsSectionCompleted = True
    
End Function

' 節の進行により発生する、あらかじめ予定されているイベントを出力
Function MakeMPBNewsSeasonEvent()
    
    Dim mpbNewsSeasonEventFlg As Boolean
    Dim mpbNewsSeasonEvent As String
    
    mpbNewsSeasonEventFlg = False
    mpbNewsSeasonEvent = "【MPB運営からのお知らせ】"

    If section = 10 Or section = 20 Then
        mpbNewsSeasonEventFlg = True
        mpbNewsSeasonEvent = AddRowText(mpbNewsSeasonEvent, "・TSOB枠の振り直しを行います。TSOB枠の表示設定を最新化してください。")
        mpbNewsSeasonEvent = AddRowText(mpbNewsSeasonEvent, "- - - - - - - - - -")
        mpbNewsSeasonEvent = AddRowText(mpbNewsSeasonEvent, "1位: " & Left(Sheets(season & "_各種記録").Cells(2, "B").Value, 1) & " → 3.5")
        mpbNewsSeasonEvent = AddRowText(mpbNewsSeasonEvent, "2位: " & Left(Sheets(season & "_各種記録").Cells(3, "B").Value, 1) & " → 4.0")
        mpbNewsSeasonEvent = AddRowText(mpbNewsSeasonEvent, "3位: " & Left(Sheets(season & "_各種記録").Cells(4, "B").Value, 1) & " → 4.5")
        mpbNewsSeasonEvent = AddRowText(mpbNewsSeasonEvent, "4位: " & Left(Sheets(season & "_各種記録").Cells(5, "B").Value, 1) & " → 5.0")
        mpbNewsSeasonEvent = AddRowText(mpbNewsSeasonEvent, "5位: " & Left(Sheets(season & "_各種記録").Cells(6, "B").Value, 1) & " → 5.5")
        mpbNewsSeasonEvent = AddRowText(mpbNewsSeasonEvent, "※同率チーム発生時には、必ずしもこの通りとならない場合があります。")
        mpbNewsSeasonEvent = AddRowText(mpbNewsSeasonEvent, "")
    End If
    
    If section = 10 Or section = 20 Then
        mpbNewsSeasonEventFlg = True
        mpbNewsSeasonEvent = AddRowText(mpbNewsSeasonEvent, "・只今より、後半戦からのHDCP変更受付を開始します。第15節終了をもって締め切るので、変更したいチームは、必要に応じて申請を行ってください。変更しない場合は、特に対応不要です。")
        mpbNewsSeasonEvent = AddRowText(mpbNewsSeasonEvent, "")
    End If
    
    If section = 15 Then
        mpbNewsSeasonEventFlg = True
        mpbNewsSeasonEvent = AddRowText(mpbNewsSeasonEvent, "・只今をもちまして、後半戦に向けたHDCP変更の申請を締め切ります。HDCPの表示設定を最新化してください。")
        mpbNewsSeasonEvent = AddRowText(mpbNewsSeasonEvent, "")
    End If
    
    If section = 25 Then
        mpbNewsSeasonEventFlg = True
        mpbNewsSeasonEvent = AddRowText(mpbNewsSeasonEvent, "・只今より、B9GGノミネートオーダーの提出受付を開始します。第28節終了をもって締め切るので、各チーム、LINEグループのアルバム「" & season & "B9GGノミネート」に提出をお願いいたします。")
        mpbNewsSeasonEvent = AddRowText(mpbNewsSeasonEvent, "")
    End If
    
    If section = 26 Or section = 27 Then
        mpbNewsSeasonEventFlg = True
        mpbNewsSeasonEvent = AddRowText(mpbNewsSeasonEvent, "・B9GGノミネートオーダーの提出/変更を受付中です。未提出のチームは、第28節が終了するまでに、LINEグループのアルバム「" & season & "B9GGノミネート」への提出をお願いいたします。")
        mpbNewsSeasonEvent = AddRowText(mpbNewsSeasonEvent, "")
    End If
    
    If section = 28 Then
        mpbNewsSeasonEventFlg = True
        mpbNewsSeasonEvent = AddRowText(mpbNewsSeasonEvent, "・B9GGノミネートオーダーを提出受付中です。未提出のチームは、第28節が終了するまでに、LINEグループのアルバム「" & season & "B9GGノミネート」への提出をお願いいたします。")
        mpbNewsSeasonEvent = AddRowText(mpbNewsSeasonEvent, "")
    End If
    
    If section = 30 Then
        mpbNewsSeasonEventFlg = True
        mpbNewsSeasonEvent = AddRowText(mpbNewsSeasonEvent, "・今シーズン、予定されていた全日程が終了しました。まずは、皆さんお疲れさまでした！この後、MPBアワードを実施しますので、案内をお待ちください。")
        mpbNewsSeasonEvent = AddRowText(mpbNewsSeasonEvent, "")
    End If
    
    If mpbNewsSeasonEventFlg Then
        mpbNewsSeasonEvent = AddRowText(mpbNewsSeasonEvent, "以上")
        Call OutputText(mpbNewsSeasonEvent, MPB_WORK_DIRECTORY_PATH & "\mpbnews-seasonevent.txt")
    End If

End Function

' 節の進行により発生する、優勝マジックや自力優勝に関するイベントを出力
Function MakeMPBNewsOfThisSection()



End Function

' スペ判定・結果を出力
Function MakeMPBNewsOfAccident()



End Function

' 次節日程調整の依頼を出力
Function MakeMPBNewsOfNextGame()



End Function

' スケジュール画像を出力
Function SavePictureOfSchedule()



End Function

' 成績画像を出力
Function SavePictureOfRanking()



End Function


Sub 画像保存()

    ' エラーチェック
    If ActiveSheet.Name <> ActiveSheet.Cells(1, "A").Value & "_スケジュール" Then
        MsgBox "シート名またはA1セルのシーズン指定が不正です。"
        End
    End If

    Application.ScreenUpdating = False
    Application.Calculate
    
    Dim seasonName As String
    Dim numberOfSection As Integer
    Dim pictureRangeSchedule, pictureRangeRanking As ChartObject
    Dim pictureName As String
    Dim minFileSize As Long
    
    seasonName = ActiveSheet.Cells(1, "A").Value
    numberOfSection = WorksheetFunction.CountIf(ActiveSheet.Range("BA2:BA241"), 0) / 8
    
    pictureName = "C:\Users\TaiNo\マイドライブ\純正リアタイ部_送信待機\schedule.jpg"
    If Dir(pictureName) <> "" Then
        MsgBox "例外が発生しました（3001）"
        End
    End If
    
    Range("A" & WorksheetFunction.Max(1, numberOfSection * 8 - 6) & ":AG" & WorksheetFunction.Max(41, numberOfSection * 8 - 6 + 57)).CopyPicture '39
    Set pictureRangeSchedule = Sheets("アクシデント").ChartObjects.Add(0, 0, Range("A" & WorksheetFunction.Max(1, numberOfSection * 8 - 6) & ":AG" & WorksheetFunction.Max(41, numberOfSection * 8 - 6 + 39)).Width, Range("A" & WorksheetFunction.Max(1, numberOfSection * 8 - 6) & ":AG" & WorksheetFunction.Max(41, numberOfSection * 8 - 6 + 57)).Height)
    pictureRangeSchedule.Chart.Export pictureName
    minFileSize = FileLen(pictureName)
    
    Do Until FileLen(pictureName) > minFileSize
        pictureRangeSchedule.Chart.Paste
        pictureRangeSchedule.Chart.Export pictureName
        DoEvents
    Loop
    
    pictureRangeSchedule.Delete
    Set pictureRangeSchedule = Nothing
    
    pictureName = "C:\Users\TaiNo\マイドライブ\純正リアタイ部_送信待機\ranking.jpg"
    If Dir(pictureName) <> "" Then
        MsgBox "例外が発生しました（3002）"
        End
    End If
    
    Sheets(seasonName & "_各種記録").Range("A1:AR41").CopyPicture
    Set pictureRangeRanking = Sheets("アクシデント").ChartObjects.Add(0, 0, Sheets(seasonName & "_各種記録").Range("A1:AR41").Width, Sheets(seasonName & "_各種記録").Range("A1:AR41").Height)
    pictureRangeRanking.Chart.Export pictureName
    minFileSize = FileLen(pictureName)
    
    Do Until FileLen(pictureName) > minFileSize
        pictureRangeRanking.Chart.Paste
        pictureRangeRanking.Chart.Export pictureName
        DoEvents
    Loop
    
    pictureRangeRanking.Delete
    Set pictureRangeRanking = Nothing
    
    Open "C:\Users\TaiNo\マイドライブ\純正リアタイ部_送信待機\nextGame.txt" For Output As #2
        Print #2, "【コミッショナーより】"
        Print #2, "試合日程の調整にご協力をお願いします。"
        Print #2, ""
        Print #2, "[第" & numberOfSection + 1 & "節]"
        If ActiveSheet.Cells(8 * numberOfSection + 3, "F") <> "" Then
            Print #2, "<実施済>　" & ActiveSheet.Cells(8 * numberOfSection + 2, "C") & " " & ActiveSheet.Cells(8 * numberOfSection + 3, "D") & " - " & ActiveSheet.Cells(8 * numberOfSection + 3, "H") & " " & ActiveSheet.Cells(8 * numberOfSection + 2, "J")
        Else
            Print #2, ActiveSheet.Cells(8 * numberOfSection + 2, "C") & "(" & ActiveSheet.Cells(8 * numberOfSection + 2, "D") & ") - (" & ActiveSheet.Cells(8 * numberOfSection + 2, "H") & ")" & ActiveSheet.Cells(8 * numberOfSection + 2, "J")
        End If
        If ActiveSheet.Cells(8 * numberOfSection + 7, "F") <> "" Then
            Print #2, "<実施済>　" & ActiveSheet.Cells(8 * numberOfSection + 6, "C") & " " & ActiveSheet.Cells(8 * numberOfSection + 7, "D") & " - " & ActiveSheet.Cells(8 * numberOfSection + 7, "H") & " " & ActiveSheet.Cells(8 * numberOfSection + 6, "J")
        Else
            Print #2, ActiveSheet.Cells(8 * numberOfSection + 6, "C") & "(" & ActiveSheet.Cells(8 * numberOfSection + 6, "D") & ") - (" & ActiveSheet.Cells(8 * numberOfSection + 6, "H") & ")" & ActiveSheet.Cells(8 * numberOfSection + 6, "J")
        End If
        Print #2, ""
        If numberOfSection < 29 Then
            Print #2, "[第" & numberOfSection + 2 & "節]"
            Print #2, ActiveSheet.Cells(8 * numberOfSection + 10, "C") & "(" & ActiveSheet.Cells(8 * numberOfSection + 10, "D") & ") - (" & ActiveSheet.Cells(8 * numberOfSection + 10, "H") & ")" & ActiveSheet.Cells(8 * numberOfSection + 10, "J")
            Print #2, ActiveSheet.Cells(8 * numberOfSection + 14, "C") & "(" & ActiveSheet.Cells(8 * numberOfSection + 14, "D") & ") - (" & ActiveSheet.Cells(8 * numberOfSection + 14, "H") & ")" & ActiveSheet.Cells(8 * numberOfSection + 14, "J");
        End If
    Close #2
    
    Call バックアップ
    
    Application.ScreenUpdating = True

End Sub

