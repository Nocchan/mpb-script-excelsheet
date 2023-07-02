Attribute VB_Name = "match_completion"
Option Explicit

Dim season As String
Dim game As Integer
Dim section As Integer

Dim dictTeamID As New Dictionary

Sub MatchCompletion()

    ' デバッグモード
    Call enableDebugMode

    ' 呼出元確認
    If Not isScheduleSheet() Then
        Call showMessageError("呼出元確認エラー", "MatchCompletion")
        End
    End If

    Call initialize

    If isSectionCompleted() Then
        Call makeMPBNewsSeasonEvent
        Call makeMPBNewsOfThisSection
        Call makeMPBNewsOfAccident
    End If

    Call makeMPBNewsOfNextGame
    Call savePictureOfSchedule
    Call savePictureOfRecord

    Call exitProcess

End Sub

' 定数・シート状態の初期化
Function initialize()

    If Not isDebugMode Then
        Application.ScreenUpdating = False
    End If

    Application.Calculate

    Call makeBackupFile

    season = ActiveSheet.Cells(1, "A").Value
    game = WorksheetFunction.CountIf(ActiveSheet.Range("F2:F241"), "（試合終了）")
    section = WorksheetFunction.CountIf(ActiveSheet.Range("BA2:BA241"), 0) / 8

    Dim teamID As Integer
    For teamID = 1 To 5
        dictTeamID.Add teamID, Sheets(season & "_各種記録").Cells(teamID + 1, "R").Value
    Next teamID

    Call Definition

    Sheets(season & "_スケジュール").Unprotect
    Sheets(season & "_投手データ").Unprotect
    Sheets(season & "_野手データ").Unprotect

End Function

' 終了時処理
Function exitProcess()

    Sheets(season & "_スケジュール").Protect AllowFormattingColumns:=True, AllowFormattingRows:=True
    Sheets(season & "_投手データ").Protect AllowFormattingColumns:=True, AllowFormattingRows:=True
    Sheets(season & "_野手データ").Protect AllowFormattingColumns:=True, AllowFormattingRows:=True

    If Not isDebugMode Then
        Application.ScreenUpdating = True
    End If

    Application.Calculate

    End

End Function

' 節が完了してスペ判定を行える状態かを判定
Function isSectionCompleted() As Boolean

    ' 消化試合数から節が完了していないことがわかるパターン
    If game <> section * 2 Then
        isSectionCompleted = False
        Exit Function
    End If

    ' 節は完了しているが不正入力があるパターン
    If Sheets(season & "_スケジュール").Cells(section * 8 + 3, "D").Value <> "" Or Sheets(season & "_スケジュール").Cells(section * 8 + 7, "D").Value <> "" Or _
       Sheets(season & "_スケジュール").Cells(section * 8 + 3, "F").Value <> "" Or Sheets(season & "_スケジュール").Cells(section * 8 + 7, "F").Value <> "" Or _
       Sheets(season & "_スケジュール").Cells(section * 8 + 3, "H").Value <> "" Or Sheets(season & "_スケジュール").Cells(section * 8 + 7, "H").Value <> "" Then
        Call showMessageError("不正入力エラー", "isSectionCompleted")
        Call exitProcess
    End If

    ' 開幕前または最終節後で予告先発を考える必要がないパターン
    If section = 0 Or section = 30 Then
        isSectionCompleted = True
        Exit Function
    End If

    ' 予告先発が出揃っていないパターン
    If Sheets(season & "_スケジュール").Cells(section * 8 + 2, "D").Value = "" Or Sheets(season & "_スケジュール").Cells(section * 8 + 6, "D").Value = "" Or _
       Sheets(season & "_スケジュール").Cells(section * 8 + 2, "H").Value = "" Or Sheets(season & "_スケジュール").Cells(section * 8 + 6, "H").Value = "" Then
        Call showMessageError("予告先発未完了エラー", "isSectionCompleted")
        Call exitProcess
    End If

    isSectionCompleted = True

End Function

' 節の進行により発生する、あらかじめ予定されているイベントを出力
Function makeMPBNewsSeasonEvent()

    ' 宣言
    Dim existMPBNewsSeasonEvent As Boolean
    Dim bodyMPBNewsSeasonEvent As String
    Dim tsobBorderDict As New Dictionary

    ' 初期化
    existMPBNewsSeasonEvent = False
    bodyMPBNewsSeasonEvent = "【MPB運営からのお知らせ】"

    ' TSOB枠の振り直し
    If section = 10 Or section = 20 Then
        existMPBNewsSeasonEvent = True

        bodyMPBNewsSeasonEvent = addLineToText(bodyMPBNewsSeasonEvent, "")
        bodyMPBNewsSeasonEvent = addLineToText(bodyMPBNewsSeasonEvent, "・TSOB枠の振り直しを行います。")
        bodyMPBNewsSeasonEvent = addLineToText(bodyMPBNewsSeasonEvent, "1位: " & Left(Sheets(season & "_各種記録").Cells(2, "B").Value, 1) & " → 3.5")
        bodyMPBNewsSeasonEvent = addLineToText(bodyMPBNewsSeasonEvent, "2位: " & Left(Sheets(season & "_各種記録").Cells(3, "B").Value, 1) & " → 4.0")
        bodyMPBNewsSeasonEvent = addLineToText(bodyMPBNewsSeasonEvent, "3位: " & Left(Sheets(season & "_各種記録").Cells(4, "B").Value, 1) & " → 4.5")
        bodyMPBNewsSeasonEvent = addLineToText(bodyMPBNewsSeasonEvent, "4位: " & Left(Sheets(season & "_各種記録").Cells(5, "B").Value, 1) & " → 5.0")
        bodyMPBNewsSeasonEvent = addLineToText(bodyMPBNewsSeasonEvent, "5位: " & Left(Sheets(season & "_各種記録").Cells(6, "B").Value, 1) & " → 5.5")
        bodyMPBNewsSeasonEvent = addLineToText(bodyMPBNewsSeasonEvent, "※同率チーム発生時には、必ずしもこの通りとならない場合があります。")

        tsobBorderDict.Add Left(Sheets(season & "_各種記録").Cells(2, "B").Value, 1), "3.5"
        tsobBorderDict.Add Left(Sheets(season & "_各種記録").Cells(3, "B").Value, 1), "4.0"
        tsobBorderDict.Add Left(Sheets(season & "_各種記録").Cells(4, "B").Value, 1), "4.5"
        tsobBorderDict.Add Left(Sheets(season & "_各種記録").Cells(5, "B").Value, 1), "5.0"
        tsobBorderDict.Add Left(Sheets(season & "_各種記録").Cells(6, "B").Value, 1), "5.5"

        Call showMessageDebug(tsobBorderDict.Item(Sheets(season & "_スケジュール").Cells(1, "BB").Value), "INPUT スケジュール.Cells(27,CP)")
        Sheets(season & "_スケジュール").Cells(27, "CP").Value = tsobBorderDict.Item(Sheets(season & "_スケジュール").Cells(1, "BB").Value)
        Call showMessageDebug(tsobBorderDict.Item(Sheets(season & "_スケジュール").Cells(1, "BC").Value), "INPUT スケジュール.Cells(27,CQ)")
        Sheets(season & "_スケジュール").Cells(27, "CQ").Value = tsobBorderDict.Item(Sheets(season & "_スケジュール").Cells(1, "BC").Value)
        Call showMessageDebug(tsobBorderDict.Item(Sheets(season & "_スケジュール").Cells(1, "BD").Value), "INPUT スケジュール.Cells(27,CR)")
        Sheets(season & "_スケジュール").Cells(27, "CR").Value = tsobBorderDict.Item(Sheets(season & "_スケジュール").Cells(1, "BD").Value)
        Call showMessageDebug(tsobBorderDict.Item(Sheets(season & "_スケジュール").Cells(1, "BE").Value), "INPUT スケジュール.Cells(27,CS)")
        Sheets(season & "_スケジュール").Cells(27, "CS").Value = tsobBorderDict.Item(Sheets(season & "_スケジュール").Cells(1, "BE").Value)
        Call showMessageDebug(tsobBorderDict.Item(Sheets(season & "_スケジュール").Cells(1, "BF").Value), "INPUT スケジュール.Cells(27,CT)")
        Sheets(season & "_スケジュール").Cells(27, "CT").Value = tsobBorderDict.Item(Sheets(season & "_スケジュール").Cells(1, "BF").Value)
    End If

    ' HDCP変更受付開始
    If section = 10 Or section = 20 Then
        existMPBNewsSeasonEvent = True

        bodyMPBNewsSeasonEvent = addLineToText(bodyMPBNewsSeasonEvent, "")
        bodyMPBNewsSeasonEvent = addLineToText(bodyMPBNewsSeasonEvent, "・只今より、後半戦からのHDCP変更受付を開始します。第15節終了をもって締め切るので、変更したいチームは、必要に応じて申請を行ってください。変更しない場合は、特に対応不要です。")
    End If

    ' HDCP変更中
    If section = 11 Or section = 12 Or section = 13 Or section = 14 Then
        existMPBNewsSeasonEvent = True

        bodyMPBNewsSeasonEvent = addLineToText(bodyMPBNewsSeasonEvent, "")
        bodyMPBNewsSeasonEvent = addLineToText(bodyMPBNewsSeasonEvent, "・後半戦からのHDCP変更を受付中です。変更したいチームは、第15節終了までに申請を行ってください。")
    End If

    ' HDCP変更受付〆
    If section = 15 Then
        existMPBNewsSeasonEvent = True

        bodyMPBNewsSeasonEvent = addLineToText(bodyMPBNewsSeasonEvent, "")
        bodyMPBNewsSeasonEvent = addLineToText(bodyMPBNewsSeasonEvent, "・只今をもちまして、後半戦に向けたHDCP変更の申請を締め切ります。HDCPの表示設定を最新化してください。")
    End If

    ' B9GG提出受付開始
    If section = 25 Then
        existMPBNewsSeasonEvent = True

        bodyMPBNewsSeasonEvent = addLineToText(bodyMPBNewsSeasonEvent, "")
        bodyMPBNewsSeasonEvent = addLineToText(bodyMPBNewsSeasonEvent, "・只今より、B9GGノミネートオーダーの提出受付を開始します。第28節終了をもって締め切るので、各チーム、LINEグループのアルバム「" & season & "B9GGノミネート」に提出をお願いいたします。")
    End If

    ' B9GG提出受付中
    If section = 26 Or section = 27 Then
        existMPBNewsSeasonEvent = True

        bodyMPBNewsSeasonEvent = addLineToText(bodyMPBNewsSeasonEvent, "")
        bodyMPBNewsSeasonEvent = addLineToText(bodyMPBNewsSeasonEvent, "・B9GGノミネートオーダーの提出/変更を受付中です。未提出のチームは、第28節が終了するまでに、LINEグループのアルバム「" & season & "B9GGノミネート」への提出をお願いいたします。")
    End If

    ' B9GG提出受付〆
    If section = 28 Then
        existMPBNewsSeasonEvent = True

        bodyMPBNewsSeasonEvent = addLineToText(bodyMPBNewsSeasonEvent, "")
        bodyMPBNewsSeasonEvent = addLineToText(bodyMPBNewsSeasonEvent, "・只今をもちまして、B9GGノミネートオーダーの提出を締め切ります。")
    End If

    ' MPBアワード案内
    If section = 30 Then
        existMPBNewsSeasonEvent = True

        bodyMPBNewsSeasonEvent = addLineToText(bodyMPBNewsSeasonEvent, "")
        bodyMPBNewsSeasonEvent = addLineToText(bodyMPBNewsSeasonEvent, "・今シーズン、予定されていた全日程が終了しました。まずは、皆さんお疲れさまでした！この後、MPBアワードを実施しますので、案内をお待ちください。")
    End If

    ' 結果の出力
    If existMPBNewsSeasonEvent Then
        bodyMPBNewsSeasonEvent = addLineToText(bodyMPBNewsSeasonEvent, "")
        bodyMPBNewsSeasonEvent = addLineToText(bodyMPBNewsSeasonEvent, "以上")

        If Not isDebugMode Then
            Call saveTxtFile(bodyMPBNewsSeasonEvent, MPB_WORK_DIRECTORY_PATH & "\batch-min\mpbnews-seasonevent.txt")
        Else
            Call showMessageInfo(bodyMPBNewsSeasonEvent, "makeMPBNewsSeasonEvent")
            Call saveTxtFile(bodyMPBNewsSeasonEvent, LOCAL_WORK_DIRECTORY_PATH & "\batch-min\mpbnews-seasonevent.txt")
        End If
    End If

End Function

' 節の進行により発生する、優勝マジックや自力優勝に関するイベントを出力
Function makeMPBNewsOfThisSection()

    ' 実行条件
    If section = 0 Then
        Exit Function
    End If

    ' 宣言
    Dim bodyMPBNewsOfThisSection As String
    Dim scoreOfThisSection(2, 4) As String
    Dim teamOfNextSection(2, 2) As String
    Dim seasonStatus As New Dictionary
    Dim headerOfNextGame As String
    Dim messageTemplateVictory As String
    Dim messageTemplateMagicDisappearance As String
    Dim messageTemplateSelfVictoryDisappearance As String
    Dim messageTemplateSelfVictoryReappearance As String
    Dim messageTemplateMagicAppearance As String

    ' 初期化
    bodyMPBNewsOfThisSection = "【MPBニュース】"
    messageTemplateVictory = season & "ペナントレース優勝が確定！"
    messageTemplateMagicDisappearance = "優勝マジックが消滅…"
    messageTemplateSelfVictoryDisappearance = "自力優勝が消滅…"
    messageTemplateSelfVictoryReappearance = "自力優勝が復活！"
    messageTemplateMagicAppearance = "優勝マジックが点灯！"
    
    ' 今節の試合結果
    If section > 0 Then
        scoreOfThisSection(1, 1) = Sheets(season & "_スケジュール").Cells(section * 8 - 6, "C").Value
        scoreOfThisSection(1, 2) = Sheets(season & "_スケジュール").Cells(section * 8 - 5, "D").Value
        scoreOfThisSection(1, 3) = Sheets(season & "_スケジュール").Cells(section * 8 - 5, "H").Value
        scoreOfThisSection(1, 4) = Sheets(season & "_スケジュール").Cells(section * 8 - 6, "J").Value
        scoreOfThisSection(2, 1) = Sheets(season & "_スケジュール").Cells(section * 8 - 2, "C").Value
        scoreOfThisSection(2, 2) = Sheets(season & "_スケジュール").Cells(section * 8 - 1, "D").Value
        scoreOfThisSection(2, 3) = Sheets(season & "_スケジュール").Cells(section * 8 - 1, "H").Value
        scoreOfThisSection(2, 4) = Sheets(season & "_スケジュール").Cells(section * 8 - 2, "J").Value
        bodyMPBNewsOfThisSection = addLineToText(bodyMPBNewsOfThisSection, "")
        bodyMPBNewsOfThisSection = addLineToText(bodyMPBNewsOfThisSection, "第" & section & "節の試合結果")
        bodyMPBNewsOfThisSection = addLineToText(bodyMPBNewsOfThisSection, scoreOfThisSection(1, 1) & " " & scoreOfThisSection(1, 2) & "-" & scoreOfThisSection(1, 3) & " " & scoreOfThisSection(1, 4))
        bodyMPBNewsOfThisSection = addLineToText(bodyMPBNewsOfThisSection, scoreOfThisSection(2, 1) & " " & scoreOfThisSection(2, 2) & "-" & scoreOfThisSection(2, 3) & " " & scoreOfThisSection(2, 4))
    End If

    ' 次節の予告先発
    If section < 30 Then
        teamOfNextSection(1, 1) = Sheets(season & "_スケジュール").Cells(section * 8 + 2, "C").Value
        teamOfNextSection(1, 2) = Sheets(season & "_スケジュール").Cells(section * 8 + 2, "J").Value
        teamOfNextSection(2, 1) = Sheets(season & "_スケジュール").Cells(section * 8 + 6, "C").Value
        teamOfNextSection(2, 2) = Sheets(season & "_スケジュール").Cells(section * 8 + 6, "J").Value
        bodyMPBNewsOfThisSection = addLineToText(bodyMPBNewsOfThisSection, "")
        bodyMPBNewsOfThisSection = addLineToText(bodyMPBNewsOfThisSection, "第" & section + 1 & "節の予告先発")
        bodyMPBNewsOfThisSection = addLineToText(bodyMPBNewsOfThisSection, "（" & Sheets(season & "_スケジュール").Cells(section * 8 + 2, "C").Value & "-" & Sheets(season & "_スケジュール").Cells(section * 8 + 2, "J").Value & "）" & Sheets(season & "_スケジュール").Cells(section * 8 + 2, "D").Value & "×" & Sheets(season & "_スケジュール").Cells(section * 8 + 2, "H").Value)
        bodyMPBNewsOfThisSection = addLineToText(bodyMPBNewsOfThisSection, "（" & Sheets(season & "_スケジュール").Cells(section * 8 + 6, "C").Value & "-" & Sheets(season & "_スケジュール").Cells(section * 8 + 6, "J").Value & "）" & Sheets(season & "_スケジュール").Cells(section * 8 + 6, "D").Value & "×" & Sheets(season & "_スケジュール").Cells(section * 8 + 6, "H").Value)
        bodyMPBNewsOfThisSection = addLineToText(bodyMPBNewsOfThisSection, "")
    End If

    ' 状況確認(今節実施前)
    seasonStatus.Add "今節実施前", seasonStatusOfSection(section, "", "", "", "", "", "")

    ' 今節実施前に優勝が決まっていない前提
    If seasonStatus.Item("今節実施前")(0) = "" Then

        seasonStatus.Add "今節実施後", seasonStatusOfSection(section, scoreOfThisSection(1, 2), Sheets(season & "_スケジュール").Cells(section * 8 - 5, "F").Value, scoreOfThisSection(1, 3), scoreOfThisSection(2, 2), Sheets(season & "_スケジュール").Cells(section * 8 - 1, "F").Value, scoreOfThisSection(2, 3))
    
        ' 今節で優勝が決まった場合
        If seasonStatus.Item("今節実施後")(0) <> "" Then
            bodyMPBNewsOfThisSection = addLineToText(bodyMPBNewsOfThisSection, "◇" & seasonStatus.Item("今節実施後")(0) & "◇" & season & "ペナントレース優勝が確定！")
        Else
            Dim teamID As Integer
            seasonStatus.Add "次節AA", seasonStatusOfSection(section + 1, "X", "tmp", "0", "X", "tmp", "0")
            seasonStatus.Add "次節BA", seasonStatusOfSection(section + 1, "0", "tmp", "X", "X", "tmp", "0")
            seasonStatus.Add "次節AB", seasonStatusOfSection(section + 1, "X", "tmp", "0", "0", "tmp", "X")
            seasonStatus.Add "次節BB", seasonStatusOfSection(section + 1, "0", "tmp", "X", "0", "tmp", "X")
            
            ' 今節
            headerOfNextGame = ""
            ' マジックが消滅するケース
            For teamID = 1 To 5
                If Left(seasonStatus.Item("今節実施前")(teamID), 1) = "M" And Left(seasonStatus.Item("今節実施後")(teamID), 1) <> "M" Then
                    bodyMPBNewsOfThisSection = addLineToText(bodyMPBNewsOfThisSection, "◇" & DICT_TEAM_NAME.Item(dictTeamID.Item(teamID)) & "◇" & headerOfNextGame & messageTemplateMagicDisappearance)
                End If
            Next teamID
            
            ' 自力優勝が消滅するケース
            For teamID = 1 To 5
                If Left(seasonStatus.Item("今節実施前")(teamID), 1) <> "自" And Left(seasonStatus.Item("今節実施後")(teamID), 1) = "自" Then
                    bodyMPBNewsOfThisSection = addLineToText(bodyMPBNewsOfThisSection, "◇" & DICT_TEAM_NAME.Item(dictTeamID.Item(teamID)) & "◇" & headerOfNextGame & messageTemplateSelfVictoryDisappearance)
                End If
            Next teamID
            
            ' 自力優勝が復活するケース
            For teamID = 1 To 5
                If Left(seasonStatus.Item("今節実施前")(teamID), 1) = "自" And Left(seasonStatus.Item("今節実施後")(teamID), 1) <> "自" Then
                    bodyMPBNewsOfThisSection = addLineToText(bodyMPBNewsOfThisSection, "◇" & DICT_TEAM_NAME.Item(dictTeamID.Item(teamID)) & "◇" & headerOfNextGame & messageTemplateSelfVictoryReappearance)
                End If
            Next teamID
            
            ' マジックが点灯するケース
            For teamID = 1 To 5
                If Left(seasonStatus.Item("今節実施前")(teamID), 1) <> "M" And Left(seasonStatus.Item("今節実施後")(teamID), 1) = "M" Then
                    bodyMPBNewsOfThisSection = addLineToText(bodyMPBNewsOfThisSection, "◇" & DICT_TEAM_NAME.Item(dictTeamID.Item(teamID)) & "◇" & headerOfNextGame & messageTemplateMagicAppearance)
                End If
            Next teamID
            
            ' 次節A_
            headerOfNextGame = "次節 " & teamOfNextSection(1, 1) & "◯-●" & teamOfNextSection(1, 2) & " で、"
            If seasonStatus.Item("次節AA")(0) <> "" And seasonStatus.Item("次節AB")(0) <> "" And seasonStatus.Item("次節AA")(0) = seasonStatus.Item("次節AB")(0) Then
                ' 優勝チームが決まるケース
                bodyMPBNewsOfThisSection = addLineToText(bodyMPBNewsOfThisSection, "◇" & seasonStatus.Item("次節AA")(0) & "◇" & headerOfNextGame & messageTemplateVictory)
            Else
                ' マジックが消滅するケース
                For teamID = 1 To 5
                    If Left(seasonStatus.Item("今節実施後")(teamID), 1) = "M" And Left(seasonStatus.Item("次節AA")(teamID), 1) <> "M" And Left(seasonStatus.Item("次節AB")(teamID), 1) <> "M" Then
                        bodyMPBNewsOfThisSection = addLineToText(bodyMPBNewsOfThisSection, "◇" & DICT_TEAM_NAME.Item(dictTeamID.Item(teamID)) & "◇" & headerOfNextGame & messageTemplateMagicDisappearance)
                    End If
                Next teamID
                
                ' 自力優勝が消滅するケース
                For teamID = 1 To 5
                    If Left(seasonStatus.Item("今節実施後")(teamID), 1) <> "自" And Left(seasonStatus.Item("次節AA")(teamID), 1) = "自" And Left(seasonStatus.Item("次節AB")(teamID), 1) = "自" Then
                        bodyMPBNewsOfThisSection = addLineToText(bodyMPBNewsOfThisSection, "◇" & DICT_TEAM_NAME.Item(dictTeamID.Item(teamID)) & "◇" & headerOfNextGame & messageTemplateSelfVictoryDisappearance)
                    End If
                Next teamID
                
                ' 自力優勝が復活するケース
                For teamID = 1 To 5
                    If Left(seasonStatus.Item("今節実施後")(teamID), 1) = "自" And Left(seasonStatus.Item("次節AA")(teamID), 1) <> "自" And Left(seasonStatus.Item("次節AB")(teamID), 1) <> "自" Then
                        bodyMPBNewsOfThisSection = addLineToText(bodyMPBNewsOfThisSection, "◇" & DICT_TEAM_NAME.Item(dictTeamID.Item(teamID)) & "◇" & headerOfNextGame & messageTemplateSelfVictoryReappearance)
                    End If
                Next teamID
                
                ' マジックが点灯するケース
                For teamID = 1 To 5
                    If Left(seasonStatus.Item("今節実施後")(teamID), 1) <> "M" And Left(seasonStatus.Item("次節AA")(teamID), 1) = "M" And Left(seasonStatus.Item("次節AB")(teamID), 1) = "M" Then
                        bodyMPBNewsOfThisSection = addLineToText(bodyMPBNewsOfThisSection, "◇" & DICT_TEAM_NAME.Item(dictTeamID.Item(teamID)) & "◇" & headerOfNextGame & messageTemplateMagicAppearance)
                    End If
                Next teamID
            End If
            
            ' 次節B_
            headerOfNextGame = "次節 " & teamOfNextSection(1, 1) & "●-◯" & teamOfNextSection(1, 2) & " で、"
            If seasonStatus.Item("次節BA")(0) <> "" And seasonStatus.Item("次節BB")(0) <> "" And seasonStatus.Item("次節BA")(0) = seasonStatus.Item("次節BB")(0) Then
                ' 優勝チームが決まるケース
                bodyMPBNewsOfThisSection = addLineToText(bodyMPBNewsOfThisSection, "◇" & seasonStatus.Item("次節BB")(0) & "◇" & headerOfNextGame & messageTemplateVictory)
            Else
                ' マジックが消滅するケース
                For teamID = 1 To 5
                    If Left(seasonStatus.Item("今節実施後")(teamID), 1) = "M" And Left(seasonStatus.Item("次節BA")(teamID), 1) <> "M" And Left(seasonStatus.Item("次節BB")(teamID), 1) <> "M" Then
                        bodyMPBNewsOfThisSection = addLineToText(bodyMPBNewsOfThisSection, "◇" & DICT_TEAM_NAME.Item(dictTeamID.Item(teamID)) & "◇" & headerOfNextGame & messageTemplateMagicDisappearance)
                    End If
                Next teamID
                
                ' 自力優勝が消滅するケース
                For teamID = 1 To 5
                    If Left(seasonStatus.Item("今節実施後")(teamID), 1) <> "自" And Left(seasonStatus.Item("次節BA")(teamID), 1) = "自" And Left(seasonStatus.Item("次節BB")(teamID), 1) = "自" Then
                        bodyMPBNewsOfThisSection = addLineToText(bodyMPBNewsOfThisSection, "◇" & DICT_TEAM_NAME.Item(dictTeamID.Item(teamID)) & "◇" & headerOfNextGame & messageTemplateSelfVictoryDisappearance)
                    End If
                Next teamID
                
                ' 自力優勝が復活するケース
                For teamID = 1 To 5
                    If Left(seasonStatus.Item("今節実施後")(teamID), 1) = "自" And Left(seasonStatus.Item("次節BA")(teamID), 1) <> "自" And Left(seasonStatus.Item("次節BB")(teamID), 1) <> "自" Then
                        bodyMPBNewsOfThisSection = addLineToText(bodyMPBNewsOfThisSection, "◇" & DICT_TEAM_NAME.Item(dictTeamID.Item(teamID)) & "◇" & headerOfNextGame & messageTemplateSelfVictoryReappearance)
                    End If
                Next teamID
                
                ' マジックが点灯するケース
                For teamID = 1 To 5
                    If Left(seasonStatus.Item("今節実施後")(teamID), 1) <> "M" And Left(seasonStatus.Item("次節BA")(teamID), 1) = "M" And Left(seasonStatus.Item("次節BB")(teamID), 1) = "M" Then
                        bodyMPBNewsOfThisSection = addLineToText(bodyMPBNewsOfThisSection, "◇" & DICT_TEAM_NAME.Item(dictTeamID.Item(teamID)) & "◇" & headerOfNextGame & messageTemplateMagicAppearance)
                    End If
                Next teamID
            End If
            
            ' 次節_A
            headerOfNextGame = "次節 " & teamOfNextSection(2, 1) & "◯-●" & teamOfNextSection(2, 2) & " で、"
            If seasonStatus.Item("次節AA")(0) <> "" And seasonStatus.Item("次節BA")(0) <> "" And seasonStatus.Item("次節AA")(0) = seasonStatus.Item("次節BA")(0) Then
                ' 優勝チームが決まるケース
                bodyMPBNewsOfThisSection = addLineToText(bodyMPBNewsOfThisSection, "◇" & seasonStatus.Item("次節BA")(0) & "◇" & headerOfNextGame & messageTemplateVictory)
            Else
                ' マジックが消滅するケース
                For teamID = 1 To 5
                    If Left(seasonStatus.Item("今節実施後")(teamID), 1) = "M" And Left(seasonStatus.Item("次節AA")(teamID), 1) <> "M" And Left(seasonStatus.Item("次節BA")(teamID), 1) <> "M" Then
                        bodyMPBNewsOfThisSection = addLineToText(bodyMPBNewsOfThisSection, "◇" & DICT_TEAM_NAME.Item(dictTeamID.Item(teamID)) & "◇" & headerOfNextGame & messageTemplateMagicDisappearance)
                    End If
                Next teamID
                
                ' 自力優勝が消滅するケース
                For teamID = 1 To 5
                    If Left(seasonStatus.Item("今節実施後")(teamID), 1) <> "自" And Left(seasonStatus.Item("次節AA")(teamID), 1) = "自" And Left(seasonStatus.Item("次節BA")(teamID), 1) = "自" Then
                        bodyMPBNewsOfThisSection = addLineToText(bodyMPBNewsOfThisSection, "◇" & DICT_TEAM_NAME.Item(dictTeamID.Item(teamID)) & "◇" & headerOfNextGame & messageTemplateSelfVictoryDisappearance)
                    End If
                Next teamID
                
                ' 自力優勝が復活するケース
                For teamID = 1 To 5
                    If Left(seasonStatus.Item("今節実施後")(teamID), 1) = "自" And Left(seasonStatus.Item("次節AA")(teamID), 1) <> "自" And Left(seasonStatus.Item("次節BA")(teamID), 1) <> "自" Then
                        bodyMPBNewsOfThisSection = addLineToText(bodyMPBNewsOfThisSection, "◇" & DICT_TEAM_NAME.Item(dictTeamID.Item(teamID)) & "◇" & headerOfNextGame & messageTemplateSelfVictoryReappearance)
                    End If
                Next teamID
                
                ' マジックが点灯するケース
                For teamID = 1 To 5
                    If Left(seasonStatus.Item("今節実施後")(teamID), 1) <> "M" And Left(seasonStatus.Item("次節AA")(teamID), 1) = "M" And Left(seasonStatus.Item("次節BA")(teamID), 1) = "M" Then
                        bodyMPBNewsOfThisSection = addLineToText(bodyMPBNewsOfThisSection, "◇" & DICT_TEAM_NAME.Item(dictTeamID.Item(teamID)) & "◇" & headerOfNextGame & messageTemplateMagicAppearance)
                    End If
                Next teamID
            End If
            
            ' 次節_B
            headerOfNextGame = "次節 " & teamOfNextSection(2, 1) & "●-◯" & teamOfNextSection(2, 2) & " で、"
            If seasonStatus.Item("次節AB")(0) <> "" And seasonStatus.Item("次節BB")(0) <> "" And seasonStatus.Item("次節AB")(0) = seasonStatus.Item("次節BB")(0) Then
                ' 優勝チームが決まるケース
                bodyMPBNewsOfThisSection = addLineToText(bodyMPBNewsOfThisSection, "◇" & seasonStatus.Item("次節AB")(0) & "◇" & headerOfNextGame & messageTemplateVictory)
            Else
                ' マジックが消滅するケース
                For teamID = 1 To 5
                    If Left(seasonStatus.Item("今節実施後")(teamID), 1) = "M" And Left(seasonStatus.Item("次節AB")(teamID), 1) <> "M" And Left(seasonStatus.Item("次節BB")(teamID), 1) <> "M" Then
                        bodyMPBNewsOfThisSection = addLineToText(bodyMPBNewsOfThisSection, "◇" & DICT_TEAM_NAME.Item(dictTeamID.Item(teamID)) & "◇" & headerOfNextGame & messageTemplateMagicDisappearance)
                    End If
                Next teamID
                
                ' 自力優勝が消滅するケース
                For teamID = 1 To 5
                    If Left(seasonStatus.Item("今節実施後")(teamID), 1) <> "自" And Left(seasonStatus.Item("次節AB")(teamID), 1) = "自" And Left(seasonStatus.Item("次節BB")(teamID), 1) = "自" Then
                        bodyMPBNewsOfThisSection = addLineToText(bodyMPBNewsOfThisSection, "◇" & DICT_TEAM_NAME.Item(dictTeamID.Item(teamID)) & "◇" & headerOfNextGame & messageTemplateSelfVictoryDisappearance)
                    End If
                Next teamID
                
                ' 自力優勝が復活するケース
                For teamID = 1 To 5
                    If Left(seasonStatus.Item("今節実施後")(teamID), 1) = "自" And Left(seasonStatus.Item("次節AB")(teamID), 1) <> "自" And Left(seasonStatus.Item("次節BB")(teamID), 1) <> "自" Then
                        bodyMPBNewsOfThisSection = addLineToText(bodyMPBNewsOfThisSection, "◇" & DICT_TEAM_NAME.Item(dictTeamID.Item(teamID)) & "◇" & headerOfNextGame & messageTemplateSelfVictoryReappearance)
                    End If
                Next teamID
                
                ' マジックが点灯するケース
                For teamID = 1 To 5
                    If Left(seasonStatus.Item("今節実施後")(teamID), 1) <> "M" And Left(seasonStatus.Item("次節AB")(teamID), 1) = "M" And Left(seasonStatus.Item("次節BB")(teamID), 1) = "M" Then
                        bodyMPBNewsOfThisSection = addLineToText(bodyMPBNewsOfThisSection, "◇" & DICT_TEAM_NAME.Item(dictTeamID.Item(teamID)) & "◇" & headerOfNextGame & messageTemplateMagicAppearance)
                    End If
                Next teamID
            End If
            
            ' 次節AA
            headerOfNextGame = "次節 " & teamOfNextSection(1, 1) & "◯-●" & teamOfNextSection(1, 2) & " & " & teamOfNextSection(2, 1) & "◯-●" & teamOfNextSection(2, 2) & " で、"
            If seasonStatus.Item("次節AA")(0) <> "" And seasonStatus.Item("次節AB")(0) <> seasonStatus.Item("次節AA")(0) And seasonStatus.Item("次節BA")(0) <> seasonStatus.Item("次節AA")(0) Then
                ' 優勝チームが決まるケース
                bodyMPBNewsOfThisSection = addLineToText(bodyMPBNewsOfThisSection, "◇" & seasonStatus.Item("次節AA")(0) & "◇" & headerOfNextGame & messageTemplateVictory)
            Else
                ' マジックが消滅するケース
                For teamID = 1 To 5
                    If Left(seasonStatus.Item("今節実施後")(teamID), 1) = "M" And Left(seasonStatus.Item("次節AA")(teamID), 1) <> "M" And Left(seasonStatus.Item("次節AB")(teamID), 1) = "M" And Left(seasonStatus.Item("次節BA")(teamID), 1) = "M" Then
                        bodyMPBNewsOfThisSection = addLineToText(bodyMPBNewsOfThisSection, "◇" & DICT_TEAM_NAME.Item(dictTeamID.Item(teamID)) & "◇" & headerOfNextGame & messageTemplateMagicDisappearance)
                    End If
                Next teamID
                
                ' 自力優勝が消滅するケース
                For teamID = 1 To 5
                    If Left(seasonStatus.Item("今節実施後")(teamID), 1) <> "自" And Left(seasonStatus.Item("次節AA")(teamID), 1) = "自" And Left(seasonStatus.Item("次節AB")(teamID), 1) <> "自" And Left(seasonStatus.Item("次節BA")(teamID), 1) <> "自" Then
                        bodyMPBNewsOfThisSection = addLineToText(bodyMPBNewsOfThisSection, "◇" & DICT_TEAM_NAME.Item(dictTeamID.Item(teamID)) & "◇" & headerOfNextGame & messageTemplateSelfVictoryDisappearance)
                    End If
                Next teamID
                
                ' 自力優勝が復活するケース
                For teamID = 1 To 5
                    If Left(seasonStatus.Item("今節実施後")(teamID), 1) = "自" And Left(seasonStatus.Item("次節AA")(teamID), 1) <> "自" And Left(seasonStatus.Item("次節AB")(teamID), 1) = "自" And Left(seasonStatus.Item("次節BA")(teamID), 1) = "自" Then
                        bodyMPBNewsOfThisSection = addLineToText(bodyMPBNewsOfThisSection, "◇" & DICT_TEAM_NAME.Item(dictTeamID.Item(teamID)) & "◇" & headerOfNextGame & messageTemplateSelfVictoryReappearance)
                    End If
                Next teamID
                
                ' マジックが点灯するケース
                For teamID = 1 To 5
                    If Left(seasonStatus.Item("今節実施後")(teamID), 1) <> "M" And Left(seasonStatus.Item("次節AA")(teamID), 1) = "M" And Left(seasonStatus.Item("次節AB")(teamID), 1) <> "M" And Left(seasonStatus.Item("次節BA")(teamID), 1) <> "M" Then
                        bodyMPBNewsOfThisSection = addLineToText(bodyMPBNewsOfThisSection, "◇" & DICT_TEAM_NAME.Item(dictTeamID.Item(teamID)) & "◇" & headerOfNextGame & messageTemplateMagicAppearance)
                    End If
                Next teamID
            End If

            ' 次節BA
            headerOfNextGame = "次節 " & teamOfNextSection(1, 1) & "●-◯" & teamOfNextSection(1, 2) & " & " & teamOfNextSection(2, 1) & "◯-●" & teamOfNextSection(2, 2) & " で、"
            If seasonStatus.Item("次節BA")(0) <> "" And seasonStatus.Item("次節BB")(0) <> seasonStatus.Item("次節BA")(0) And seasonStatus.Item("次節AA")(0) <> seasonStatus.Item("次節BA")(0) Then
                ' 優勝チームが決まるケース
                bodyMPBNewsOfThisSection = addLineToText(bodyMPBNewsOfThisSection, "◇" & seasonStatus.Item("次節BA")(0) & "◇" & headerOfNextGame & messageTemplateVictory)
            Else
                ' マジックが消滅するケース
                For teamID = 1 To 5
                    If Left(seasonStatus.Item("今節実施後")(teamID), 1) = "M" And Left(seasonStatus.Item("次節BA")(teamID), 1) <> "M" And Left(seasonStatus.Item("次節BB")(teamID), 1) = "M" And Left(seasonStatus.Item("次節AA")(teamID), 1) = "M" Then
                        bodyMPBNewsOfThisSection = addLineToText(bodyMPBNewsOfThisSection, "◇" & DICT_TEAM_NAME.Item(dictTeamID.Item(teamID)) & "◇" & headerOfNextGame & messageTemplateMagicDisappearance)
                    End If
                Next teamID
                
                ' 自力優勝が消滅するケース
                For teamID = 1 To 5
                    If Left(seasonStatus.Item("今節実施後")(teamID), 1) <> "自" And Left(seasonStatus.Item("次節BA")(teamID), 1) = "自" And Left(seasonStatus.Item("次節BB")(teamID), 1) <> "自" And Left(seasonStatus.Item("次節AA")(teamID), 1) <> "自" Then
                        bodyMPBNewsOfThisSection = addLineToText(bodyMPBNewsOfThisSection, "◇" & DICT_TEAM_NAME.Item(dictTeamID.Item(teamID)) & "◇" & headerOfNextGame & messageTemplateSelfVictoryDisappearance)
                    End If
                Next teamID
                
                ' 自力優勝が復活するケース
                For teamID = 1 To 5
                    If Left(seasonStatus.Item("今節実施後")(teamID), 1) = "自" And Left(seasonStatus.Item("次節BA")(teamID), 1) <> "自" And Left(seasonStatus.Item("次節BB")(teamID), 1) = "自" And Left(seasonStatus.Item("次節AA")(teamID), 1) = "自" Then
                        bodyMPBNewsOfThisSection = addLineToText(bodyMPBNewsOfThisSection, "◇" & DICT_TEAM_NAME.Item(dictTeamID.Item(teamID)) & "◇" & headerOfNextGame & messageTemplateSelfVictoryReappearance)
                    End If
                Next teamID
                
                ' マジックが点灯するケース
                For teamID = 1 To 5
                    If Left(seasonStatus.Item("今節実施後")(teamID), 1) <> "M" And Left(seasonStatus.Item("次節BA")(teamID), 1) = "M" And Left(seasonStatus.Item("次節BB")(teamID), 1) <> "M" And Left(seasonStatus.Item("次節AA")(teamID), 1) <> "M" Then
                        bodyMPBNewsOfThisSection = addLineToText(bodyMPBNewsOfThisSection, "◇" & DICT_TEAM_NAME.Item(dictTeamID.Item(teamID)) & "◇" & headerOfNextGame & messageTemplateMagicAppearance)
                    End If
                Next teamID
            End If
            
            ' 次節AB
            headerOfNextGame = "次節 " & teamOfNextSection(1, 1) & "◯-●" & teamOfNextSection(1, 2) & " & " & teamOfNextSection(2, 1) & "●-◯" & teamOfNextSection(2, 2) & " で、"
            If seasonStatus.Item("次節AB")(0) <> "" And seasonStatus.Item("次節AA")(0) <> seasonStatus.Item("次節AB")(0) And seasonStatus.Item("次節BB")(0) <> seasonStatus.Item("次節AB")(0) Then
                ' 優勝チームが決まるケース
                bodyMPBNewsOfThisSection = addLineToText(bodyMPBNewsOfThisSection, "◇" & seasonStatus.Item("次節AB")(0) & "◇" & headerOfNextGame & messageTemplateVictory)
            Else
                ' マジックが消滅するケース
                For teamID = 1 To 5
                    If Left(seasonStatus.Item("今節実施後")(teamID), 1) = "M" And Left(seasonStatus.Item("次節AB")(teamID), 1) <> "M" And Left(seasonStatus.Item("次節AA")(teamID), 1) = "M" And Left(seasonStatus.Item("次節BB")(teamID), 1) = "M" Then
                        bodyMPBNewsOfThisSection = addLineToText(bodyMPBNewsOfThisSection, "◇" & DICT_TEAM_NAME.Item(dictTeamID.Item(teamID)) & "◇" & headerOfNextGame & messageTemplateMagicDisappearance)
                    End If
                Next teamID
                
                ' 自力優勝が消滅するケース
                For teamID = 1 To 5
                    If Left(seasonStatus.Item("今節実施後")(teamID), 1) <> "自" And Left(seasonStatus.Item("次節AB")(teamID), 1) = "自" And Left(seasonStatus.Item("次節AA")(teamID), 1) <> "自" And Left(seasonStatus.Item("次節BB")(teamID), 1) <> "自" Then
                        bodyMPBNewsOfThisSection = addLineToText(bodyMPBNewsOfThisSection, "◇" & DICT_TEAM_NAME.Item(dictTeamID.Item(teamID)) & "◇" & headerOfNextGame & messageTemplateSelfVictoryDisappearance)
                    End If
                Next teamID
                
                ' 自力優勝が復活するケース
                For teamID = 1 To 5
                    If Left(seasonStatus.Item("今節実施後")(teamID), 1) = "自" And Left(seasonStatus.Item("次節AB")(teamID), 1) <> "自" And Left(seasonStatus.Item("次節AA")(teamID), 1) = "自" And Left(seasonStatus.Item("次節BB")(teamID), 1) = "自" Then
                        bodyMPBNewsOfThisSection = addLineToText(bodyMPBNewsOfThisSection, "◇" & DICT_TEAM_NAME.Item(dictTeamID.Item(teamID)) & "◇" & headerOfNextGame & messageTemplateSelfVictoryReappearance)
                    End If
                Next teamID
                
                ' マジックが点灯するケース
                For teamID = 1 To 5
                    If Left(seasonStatus.Item("今節実施後")(teamID), 1) <> "M" And Left(seasonStatus.Item("次節AB")(teamID), 1) = "M" And Left(seasonStatus.Item("次節AA")(teamID), 1) <> "M" And Left(seasonStatus.Item("次節BB")(teamID), 1) <> "M" Then
                        bodyMPBNewsOfThisSection = addLineToText(bodyMPBNewsOfThisSection, "◇" & DICT_TEAM_NAME.Item(dictTeamID.Item(teamID)) & "◇" & headerOfNextGame & messageTemplateMagicAppearance)
                    End If
                Next teamID
            End If
            
            ' 次節BB
            headerOfNextGame = "次節 " & teamOfNextSection(1, 1) & "●-◯" & teamOfNextSection(1, 2) & " & " & teamOfNextSection(2, 1) & "●-◯" & teamOfNextSection(2, 2) & " で、"
            If seasonStatus.Item("次節BB")(0) <> "" And seasonStatus.Item("次節BA")(0) <> seasonStatus.Item("次節BB")(0) And seasonStatus.Item("次節AB")(0) <> seasonStatus.Item("次節BB")(0) Then
                ' 優勝チームが決まるケース
                bodyMPBNewsOfThisSection = addLineToText(bodyMPBNewsOfThisSection, "◇" & seasonStatus.Item("次節BB")(0) & "◇" & headerOfNextGame & messageTemplateVictory)
            Else
                ' マジックが消滅するケース
                For teamID = 1 To 5
                    If Left(seasonStatus.Item("今節実施後")(teamID), 1) = "M" And Left(seasonStatus.Item("次節BB")(teamID), 1) <> "M" And Left(seasonStatus.Item("次節BA")(teamID), 1) = "M" And Left(seasonStatus.Item("次節AB")(teamID), 1) = "M" Then
                        bodyMPBNewsOfThisSection = addLineToText(bodyMPBNewsOfThisSection, "◇" & DICT_TEAM_NAME.Item(dictTeamID.Item(teamID)) & "◇" & headerOfNextGame & messageTemplateMagicDisappearance)
                    End If
                Next teamID
                
                ' 自力優勝が消滅するケース
                For teamID = 1 To 5
                    If Left(seasonStatus.Item("今節実施後")(teamID), 1) <> "自" And Left(seasonStatus.Item("次節BB")(teamID), 1) = "自" And Left(seasonStatus.Item("次節BA")(teamID), 1) <> "自" And Left(seasonStatus.Item("次節AB")(teamID), 1) <> "自" Then
                        bodyMPBNewsOfThisSection = addLineToText(bodyMPBNewsOfThisSection, "◇" & DICT_TEAM_NAME.Item(dictTeamID.Item(teamID)) & "◇" & headerOfNextGame & messageTemplateSelfVictoryDisappearance)
                    End If
                Next teamID
                
                ' 自力優勝が復活するケース
                For teamID = 1 To 5
                    If Left(seasonStatus.Item("今節実施後")(teamID), 1) = "自" And Left(seasonStatus.Item("次節BB")(teamID), 1) <> "自" And Left(seasonStatus.Item("次節BA")(teamID), 1) = "自" And Left(seasonStatus.Item("次節AB")(teamID), 1) = "自" Then
                        bodyMPBNewsOfThisSection = addLineToText(bodyMPBNewsOfThisSection, "◇" & DICT_TEAM_NAME.Item(dictTeamID.Item(teamID)) & "◇" & headerOfNextGame & messageTemplateSelfVictoryReappearance)
                    End If
                Next teamID
                
                ' マジックが点灯するケース
                For teamID = 1 To 5
                    If Left(seasonStatus.Item("今節実施後")(teamID), 1) <> "M" And Left(seasonStatus.Item("次節BB")(teamID), 1) = "M" And Left(seasonStatus.Item("次節BA")(teamID), 1) <> "M" And Left(seasonStatus.Item("次節AB")(teamID), 1) <> "M" Then
                        bodyMPBNewsOfThisSection = addLineToText(bodyMPBNewsOfThisSection, "◇" & DICT_TEAM_NAME.Item(dictTeamID.Item(teamID)) & "◇" & headerOfNextGame & messageTemplateMagicAppearance)
                    End If
                Next teamID
            End If
        End If
    End If

    ' 結果の出力
    If Not isDebugMode Then
        Call saveTxtFile(bodyMPBNewsOfThisSection, MPB_WORK_DIRECTORY_PATH & "\batch-min\mpbnews-section.txt")
    Else
        Call showMessageInfo(bodyMPBNewsOfThisSection, "makeMPBNewsOfThisSection")
        Call saveTxtFile(bodyMPBNewsOfThisSection, LOCAL_WORK_DIRECTORY_PATH & "\batch-min\mpbnews-section.txt")
    End If

End Function

Function seasonStatusOfSection(sectionNumber As Integer, score1D As String, score1F As String, score1H As String, score2D As String, score2F As String, score2H As String) As String()

    Dim tmp(2, 3) As String
    Dim resultArray(5) As String
        
    tmp(1, 1) = Sheets(season & "_スケジュール").Cells(sectionNumber * 8 - 5, "D").Value
    tmp(1, 2) = Sheets(season & "_スケジュール").Cells(sectionNumber * 8 - 5, "F").Value
    tmp(1, 3) = Sheets(season & "_スケジュール").Cells(sectionNumber * 8 - 5, "H").Value
    tmp(2, 1) = Sheets(season & "_スケジュール").Cells(sectionNumber * 8 - 1, "D").Value
    tmp(2, 2) = Sheets(season & "_スケジュール").Cells(sectionNumber * 8 - 1, "F").Value
    tmp(2, 3) = Sheets(season & "_スケジュール").Cells(sectionNumber * 8 - 1, "H").Value
    
    Sheets(season & "_スケジュール").Cells(sectionNumber * 8 - 5, "D").Value = score1D
    Sheets(season & "_スケジュール").Cells(sectionNumber * 8 - 5, "F").Value = score1F
    Sheets(season & "_スケジュール").Cells(sectionNumber * 8 - 5, "H").Value = score1H
    Sheets(season & "_スケジュール").Cells(sectionNumber * 8 - 1, "D").Value = score2D
    Sheets(season & "_スケジュール").Cells(sectionNumber * 8 - 1, "F").Value = score2F
    Sheets(season & "_スケジュール").Cells(sectionNumber * 8 - 1, "H").Value = score2H
    
    Application.Calculate

    Dim teamID As Integer
    resultArray(0) = ""
    For teamID = 1 To 5

        resultArray(teamID) = "-"

        If Sheets(season & "_各種記録").Cells(teamID + 1, "BR").Value = 0 Then
            resultArray(teamID) = "自力V消滅"
        ElseIf Sheets(season & "_各種記録").Cells(teamID + 1, "BX").Value = "優勝" Then
            resultArray(teamID) = Sheets(season & "_各種記録").Cells(teamID + 1, "BX").Value
            resultArray(0) = DICT_TEAM_NAME.Item(dictTeamID.Item(teamID))
        ElseIf Sheets(season & "_各種記録").Cells(teamID + 1, "BX").Value <> "-" Then
            resultArray(teamID) = Sheets(season & "_各種記録").Cells(teamID + 1, "BX").Value
        End If

    Next teamID
    
    Sheets(season & "_スケジュール").Cells(sectionNumber * 8 - 5, "D").Value = tmp(1, 1)
    Sheets(season & "_スケジュール").Cells(sectionNumber * 8 - 5, "F").Value = tmp(1, 2)
    Sheets(season & "_スケジュール").Cells(sectionNumber * 8 - 5, "H").Value = tmp(1, 3)
    Sheets(season & "_スケジュール").Cells(sectionNumber * 8 - 1, "D").Value = tmp(2, 1)
    Sheets(season & "_スケジュール").Cells(sectionNumber * 8 - 1, "F").Value = tmp(2, 2)
    Sheets(season & "_スケジュール").Cells(sectionNumber * 8 - 1, "H").Value = tmp(2, 3)
    
    Application.Calculate

    seasonStatusOfSection = resultArray()

End Function

' スペ判定・結果を出力
Function makeMPBNewsOfAccident()

    ' 実行条件
    If section = 30 Then
        Exit Function
    End If

    ' 宣言
    Dim existMPBNewsOfAccident As Boolean
    Dim bodyMPBNewsOfAccident As String
    Dim gamesBeforeThisSection As Integer
    Dim gamesAfterThisSection As Integer
    Dim teamBasedAccidentRate As Single

    ' 初期化
    existMPBNewsOfAccident = False
    bodyMPBNewsOfAccident = "【選手離脱情報】"
    bodyMPBNewsOfAccident = addLineToText(bodyMPBNewsOfAccident, "")
    gamesBeforeThisSection = -1
    gamesAfterThisSection = 0
    teamBasedAccidentRate = 0

    Dim teamID As Integer
    Dim rowIdx As Integer
    Dim columnIdx As Integer
    Dim dice As Single
    Dim visibleAccidentPeriod As Integer
    Dim hiddenAccidentPeriod As Integer
    Dim accidentInformation As String
    For teamID = 1 To 5

        ' 試合状況チェック
        If section > 0 Then
            gamesBeforeThisSection = Sheets(season & "_スケジュール").Cells(2 + section - 1, 83 + teamID)
        End If
        gamesAfterThisSection = Sheets(season & "_スケジュール").Cells(2 + section, 83 + teamID)

        ' 基礎スペ率=(BASE_ACCIDENT_RATE)*(ヤ戦病院適用分)*(試合進行係数88.5-111.5%) ※試合していない場合はゼロ
        teamBasedAccidentRate = BASE_ACCIDENT_RATE * DICT_ACCIDENT_HDCP.Item(dictTeamID.Item(teamID)) * (0.885 + (gamesAfterThisSection * 0.01))
        If gamesBeforeThisSection = gamesAfterThisSection Then
            teamBasedAccidentRate = 0
        End If
        ' Call showMessageInfo(dictTeamID.Item(teamID) & " : teamBasedAccidentRate = " & teamBasedAccidentRate * 100 & "%", "makeMPBNewsOfAccident")

        ' 投手スペ判定
        For rowIdx = 4 + 50 * (teamID - 1) To 50 * teamID

            If Sheets(season & "_投手データ").Cells(rowIdx, "A").Value = "" Then
                Exit For
            End If

            ' 基礎スペ率*スペ査定係数での抽選 ※既にケガしている場合は対象外
            If Sheets(season & "_投手データ").Cells(rowIdx, 282 + gamesAfterThisSection).Value = "" Then
                Randomize
                dice = Rnd()
            Else
                dice = 1#
            End If
            ' Call showMessageInfo(dictTeamID.Item(teamID) & Sheets(season & "_投手データ").Cells(rowIdx, "D").Value & " : accidentRate = " & teamBasedAccidentRate * DICT_ACCIDENT_COEFFICIENT.Item(Sheets(season & "_投手データ").Cells(rowIdx, "E").Value) * 100 & "%, dice = " & dice * 100, "makeMPBNewsOfAccident")
            If dice < teamBasedAccidentRate * DICT_ACCIDENT_COEFFICIENT.Item(Sheets(season & "_投手データ").Cells(rowIdx, "E").Value) Then

                ' スペ長さ(表)抽選
                visibleAccidentPeriod = drawFromDictionary(DICT_ACCIDENT_LENGTH_RATE)

                ' スペ長さ(裏)抽選 ※長さゼロにはならない、今期絶望の場合は変動なし
                hiddenAccidentPeriod = visibleAccidentPeriod + drawFromDictionary(DICT_ACCIDENT_MARGIN_DICT.Item(Sheets(season & "_投手データ").Cells(rowIdx, "E").Value))
                If hiddenAccidentPeriod = 0 Then
                    hiddenAccidentPeriod = 1
                End If
                If visibleAccidentPeriod = 24 Then
                    hiddenAccidentPeriod = 24
                End If

                ' スペ内容抽選
                accidentInformation = drawFromDictionary(DICT_ACCIDENT_INFORMATION_PITCHER_DICT.Item(visibleAccidentPeriod))
                bodyMPBNewsOfAccident = addLineToText(bodyMPBNewsOfAccident, "◇" & DICT_TEAM_NAME.Item(dictTeamID.Item(teamID)) & "◇" & Sheets(season & "_投手データ").Cells(rowIdx, "D").Value & accidentInformation)
                existMPBNewsOfAccident = True

                ' ファイル書き込み
                For columnIdx = 282 + gamesAfterThisSection To 282 + gamesAfterThisSection + hiddenAccidentPeriod - 1
                    If columnIdx > 305 Then
                        Exit For
                    End If
                    If visibleAccidentPeriod <> 24 Then
                        ' Call showMessageDebug(Sheets(season & "_投手データ").Cells(rowIdx, "D").Value & "(" & columnIdx - (282 + gamesAfterThisSection) + 1 & "/" & visibleAccidentPeriod & ")", "INPUT 投手データ.Cells(" & rowIdx & "," & columnIdx & ")")
                        Sheets(season & "_投手データ").Cells(rowIdx, columnIdx).Value = Sheets(season & "_投手データ").Cells(rowIdx, "D").Value & "(" & columnIdx - (282 + gamesAfterThisSection) + 1 & "/" & visibleAccidentPeriod & ")"
                    Else
                        ' Call showMessageDebug(Sheets(season & "_投手データ").Cells(rowIdx, "D").Value & "(-)", "INPUT 投手データ.Cells(" & rowIdx & "," & columnIdx & ")")
                        Sheets(season & "_投手データ").Cells(rowIdx, columnIdx).Value = Sheets(season & "_投手データ").Cells(rowIdx, "D").Value & "(-)"
                    End If
                Next columnIdx

            ElseIf Sheets(season & "_投手データ").Cells(rowIdx, 282 + gamesBeforeThisSection).Value <> "" And Sheets(season & "_投手データ").Cells(rowIdx, 282 + gamesAfterThisSection).Value = "" Then

                ' 復帰
                bodyMPBNewsOfAccident = addLineToText(bodyMPBNewsOfAccident, "◇" & DICT_TEAM_NAME.Item(dictTeamID.Item(teamID)) & "◇離脱中の" & Sheets(season & "_投手データ").Cells(rowIdx, "D").Value & "選手について、次節からの戦列復帰が明言されました。")
                existMPBNewsOfAccident = True

            End If

        Next rowIdx

        ' 野手スペ判定
        Randomize
        For rowIdx = 4 + 50 * (teamID - 1) To 50 * teamID

            If Sheets(season & "_野手データ").Cells(rowIdx, "A").Value = "" Then
                Exit For
            End If

            ' 基礎スペ率*スペ査定係数での抽選 ※既にケガしている場合は対象外
            If Sheets(season & "_野手データ").Cells(rowIdx, 236 + gamesAfterThisSection).Value = "" Then
                Randomize
                dice = Rnd()
            Else
                dice = 1
            End If
            ' Call showMessageInfo(dictTeamID.Item(teamID) & Sheets(season & "_野手データ").Cells(rowIdx, "D").Value & " : accidentRate = " & teamBasedAccidentRate * DICT_ACCIDENT_COEFFICIENT.Item(Sheets(season & "_野手データ").Cells(rowIdx, "E").Value) * 100 & "%, dice = " & dice * 100, "makeMPBNewsOfAccident")
            If dice < teamBasedAccidentRate * DICT_ACCIDENT_COEFFICIENT.Item(Sheets(season & "_野手データ").Cells(rowIdx, "E").Value) Then

                ' スペ長さ(表)抽選
                visibleAccidentPeriod = drawFromDictionary(DICT_ACCIDENT_LENGTH_RATE)

                ' スペ長さ(裏)抽選 ※長さゼロにはならない、今期絶望の場合は変動なし
                hiddenAccidentPeriod = visibleAccidentPeriod + drawFromDictionary(DICT_ACCIDENT_MARGIN_DICT.Item(Sheets(season & "_野手データ").Cells(rowIdx, "E").Value))
                If hiddenAccidentPeriod = 0 Then
                    hiddenAccidentPeriod = 1
                End If
                If visibleAccidentPeriod = 24 Then
                    hiddenAccidentPeriod = 24
                End If

                ' スペ内容抽選
                accidentInformation = drawFromDictionary(DICT_ACCIDENT_INFORMATION_FIELDER_DICT.Item(visibleAccidentPeriod))
                bodyMPBNewsOfAccident = addLineToText(bodyMPBNewsOfAccident, "◇" & DICT_TEAM_NAME.Item(dictTeamID.Item(teamID)) & "◇" & Sheets(season & "_野手データ").Cells(rowIdx, "D").Value & accidentInformation)
                existMPBNewsOfAccident = True

                ' ファイル書き込み
                For columnIdx = 236 + gamesAfterThisSection To 236 + gamesAfterThisSection + hiddenAccidentPeriod - 1
                    If columnIdx > 259 Then
                        Exit For
                    End If
                    If visibleAccidentPeriod <> 24 Then
                        ' Call showMessageDebug(Sheets(season & "_野手データ").Cells(rowIdx, "D").Value & "(" & columnIdx - (236 + gamesAfterThisSection) + 1 & "/" & visibleAccidentPeriod & ")", "INPUT 野手データ.Cells(" & rowIdx & "," & columnIdx & ")")
                        Sheets(season & "_野手データ").Cells(rowIdx, columnIdx).Value = Sheets(season & "_野手データ").Cells(rowIdx, "D").Value & "(" & columnIdx - (236 + gamesAfterThisSection) + 1 & "/" & visibleAccidentPeriod & ")"
                    Else
                        ' Call showMessageDebug(Sheets(season & "_野手データ").Cells(rowIdx, "D").Value & "(-)", "INPUT 野手データ.Cells(" & rowIdx & "," & columnIdx & ")")
                        Sheets(season & "_野手データ").Cells(rowIdx, columnIdx).Value = Sheets(season & "_野手データ").Cells(rowIdx, "D").Value & "(-)"
                    End If
                Next columnIdx

            ElseIf Sheets(season & "_野手データ").Cells(rowIdx, 236 + gamesBeforeThisSection).Value <> "" And Sheets(season & "_野手データ").Cells(rowIdx, 236 + gamesAfterThisSection).Value = "" Then

                ' 復帰
                bodyMPBNewsOfAccident = addLineToText(bodyMPBNewsOfAccident, "◇" & DICT_TEAM_NAME.Item(dictTeamID.Item(teamID)) & "◇離脱中の" & Sheets(season & "_野手データ").Cells(rowIdx, "D").Value & "選手について、次節からの戦列復帰が明言されました。")
                existMPBNewsOfAccident = True

            End If

        Next rowIdx

    Next teamID

    ' 結果の出力
    If Not existMPBNewsOfAccident Then
        bodyMPBNewsOfAccident = addLineToText(bodyMPBNewsOfAccident, "選手の離脱/復帰に関する情報はありません。")
    End If

    If Not isDebugMode Then
        Call saveTxtFile(bodyMPBNewsOfAccident, MPB_WORK_DIRECTORY_PATH & "\batch-min\mpbnews-accident.txt")
    Else
        Call showMessageInfo(bodyMPBNewsOfAccident, "makeMPBNewsOfAccident")
        Call saveTxtFile(bodyMPBNewsOfAccident, LOCAL_WORK_DIRECTORY_PATH & "\batch-min\mpbnews-accident.txt")
    End If

End Function

' 次節日程調整の依頼を出力
Function makeMPBNewsOfNextGame()

    ' 実行条件
    If section = 30 Then
        Exit Function
    End If

    ' 宣言
    Dim bodyMPBNewsOfNextGame As String

    ' 初期化
    bodyMPBNewsOfNextGame = addLineToText(bodyMPBNewsOfNextGame, "試合日程の調整にご協力をお願いいたします。")
    bodyMPBNewsOfNextGame = addLineToText(bodyMPBNewsOfNextGame, "")

    bodyMPBNewsOfNextGame = addLineToText(bodyMPBNewsOfNextGame, "[第" & section + 1 & "節]")
    If Sheets(season & "_スケジュール").Cells(8 * section + 3, "F").Value <> "" Then
        bodyMPBNewsOfNextGame = addLineToText(bodyMPBNewsOfNextGame, "<実施済>　" & Sheets(season & "_スケジュール").Cells(8 * section + 2, "C").Value & " " & Sheets(season & "_スケジュール").Cells(8 * section + 3, "D").Value & " - " & Sheets(season & "_スケジュール").Cells(8 * section + 3, "H").Value & " " & Sheets(season & "_スケジュール").Cells(8 * section + 2, "J").Value)
    Else
        bodyMPBNewsOfNextGame = addLineToText(bodyMPBNewsOfNextGame, Sheets(season & "_スケジュール").Cells(8 * section + 2, "C").Value & "(" & Sheets(season & "_スケジュール").Cells(8 * section + 2, "D").Value & ") - (" & Sheets(season & "_スケジュール").Cells(8 * section + 2, "H").Value & ") " & Sheets(season & "_スケジュール").Cells(8 * section + 2, "J").Value)
    End If
    If Sheets(season & "_スケジュール").Cells(8 * section + 7, "F").Value Then
        bodyMPBNewsOfNextGame = addLineToText(bodyMPBNewsOfNextGame, "<実施済>　" & Sheets(season & "_スケジュール").Cells(8 * section + 6, "C").Value & " " & Sheets(season & "_スケジュール").Cells(8 * section + 7, "D").Value & " - " & Sheets(season & "_スケジュール").Cells(8 * section + 7, "H").Value & " " & Sheets(season & "_スケジュール").Cells(8 * section + 6, "J").Value)
    Else
        bodyMPBNewsOfNextGame = addLineToText(bodyMPBNewsOfNextGame, Sheets(season & "_スケジュール").Cells(8 * section + 6, "C").Value & "(" & Sheets(season & "_スケジュール").Cells(8 * section + 6, "D").Value & ") - (" & Sheets(season & "_スケジュール").Cells(8 * section + 6, "H").Value & ") " & Sheets(season & "_スケジュール").Cells(8 * section + 6, "J").Value)
    End If

    bodyMPBNewsOfNextGame = addLineToText(bodyMPBNewsOfNextGame, "")

    If section <= 28 Then
        bodyMPBNewsOfNextGame = addLineToText(bodyMPBNewsOfNextGame, "[第" & section + 2 & "節]")
        bodyMPBNewsOfNextGame = addLineToText(bodyMPBNewsOfNextGame, Sheets(season & "_スケジュール").Cells(8 * section + 10, "C").Value & "(" & Sheets(season & "_スケジュール").Cells(8 * section + 10, "D").Value & ") - (" & Sheets(season & "_スケジュール").Cells(8 * section + 10, "H").Value & ") " & Sheets(season & "_スケジュール").Cells(8 * section + 10, "J").Value)
        bodyMPBNewsOfNextGame = addLineToText(bodyMPBNewsOfNextGame, Sheets(season & "_スケジュール").Cells(8 * section + 14, "C").Value & "(" & Sheets(season & "_スケジュール").Cells(8 * section + 14, "D").Value & ") - (" & Sheets(season & "_スケジュール").Cells(8 * section + 14, "H").Value & ") " & Sheets(season & "_スケジュール").Cells(8 * section + 14, "J").Value)
    End If

    If Not isDebugMode Then
        Call saveTxtFile(bodyMPBNewsOfNextGame, MPB_WORK_DIRECTORY_PATH & "\batch-week\mpbnews-nextgame.txt")
    Else
        Call showMessageInfo(bodyMPBNewsOfNextGame, "makeMPBNewsOfNextGame")
        Call saveTxtFile(bodyMPBNewsOfNextGame, LOCAL_WORK_DIRECTORY_PATH & "\batch-week\mpbnews-nextgame.txt")
    End If

End Function

' スケジュール画像を出力
Function savePictureOfSchedule()

    Application.Calculate

    If Not isDebugMode Then
        Call savePngFile(Sheets(season & "_スケジュール").Range("A" & WorksheetFunction.Max(1, section * 8 - 6) & ":AG" & WorksheetFunction.Max(41, section * 8 - 6 + 55)), MPB_WORK_DIRECTORY_PATH & "\batch-min\mpbpicture-schedule.png")
    Else
        Call savePngFile(Sheets(season & "_スケジュール").Range("A" & WorksheetFunction.Max(1, section * 8 - 6) & ":AG" & WorksheetFunction.Max(41, section * 8 - 6 + 55)), LOCAL_WORK_DIRECTORY_PATH & "\batch-min\mpbpicture-schedule.png")
    End If


End Function

' 成績画像を出力
Function savePictureOfRecord()

    Application.Calculate

    If Not isDebugMode Then
        Call savePngFile(Sheets(season & "_各種記録").Range("A1:AR41"), MPB_WORK_DIRECTORY_PATH & "\batch-min\mpbpicture-record.png")
    Else
        Call savePngFile(Sheets(season & "_各種記録").Range("A1:AR41"), LOCAL_WORK_DIRECTORY_PATH & "\batch-min\mpbpicture-record.png")
    End If

End Function
