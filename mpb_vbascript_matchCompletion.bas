Attribute VB_Name = "mpb_vbascript_matchCompletion"
Option Explicit

Dim season As String
Dim game As Integer
Dim section As Integer

Dim dictTeamID As New Dictionary

Sub matchCompletion()
    
    ' デバッグモード
    Call DebugMode

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
    
    Call Backup
    
    season = ActiveSheet.Cells(1, "A").Value
    game = WorksheetFunction.CountIf(ActiveSheet.Range("BA2:BA241"), 0) / 4
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
Function ExitProcess()
    
    Sheets(season & "_スケジュール").Protect AllowFormattingColumns:=True, AllowFormattingRows:=True
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
    If Sheets(season & "_スケジュール").Cells(section * 8 + 3, "D").Value <> "" Or Sheets(season & "_スケジュール").Cells(section * 8 + 7, "D").Value <> "" Or _
       Sheets(season & "_スケジュール").Cells(section * 8 + 3, "F").Value <> "" Or Sheets(season & "_スケジュール").Cells(section * 8 + 7, "F").Value <> "" Or _
       Sheets(season & "_スケジュール").Cells(section * 8 + 3, "H").Value <> "" Or Sheets(season & "_スケジュール").Cells(section * 8 + 7, "H").Value <> "" Then
        Call MessageError("不正入力エラー", "IsSectionCompleted")
        Call ExitProcess
    End If
    
    ' 開幕前または最終節後で予告先発を考える必要がないパターン
    If section = 0 Or section = 30 Then
        IsSectionCompleted = True
        Exit Function
    End If
    
    ' 予告先発が出揃っていないパターン
    If Sheets(season & "_スケジュール").Cells(section * 8 + 2, "D").Value = "" Or Sheets(season & "_スケジュール").Cells(section * 8 + 6, "D").Value = "" Or _
       Sheets(season & "_スケジュール").Cells(section * 8 + 2, "H").Value = "" Or Sheets(season & "_スケジュール").Cells(section * 8 + 6, "H").Value = "" Then
        Call MessageError("予告先発未完了エラー", "IsSectionCompleted")
        Call ExitProcess
    End If
    
    IsSectionCompleted = True
    
End Function

' 節の進行により発生する、あらかじめ予定されているイベントを出力
Function MakeMPBNewsSeasonEvent()
    
    ' 宣言
    Dim mpbNewsSeasonEventFlg As Boolean
    Dim mpbNewsSeasonEvent As String
    Dim tsobBorderDict As New Dictionary
    
    ' 初期化
    mpbNewsSeasonEventFlg = False
    mpbNewsSeasonEvent = "【MPB運営からのお知らせ】"
    
    ' TSOB枠の振り直し
    If section = 10 Or section = 20 Then
        mpbNewsSeasonEventFlg = True
        
        mpbNewsSeasonEvent = AddRowText(mpbNewsSeasonEvent, "・TSOB枠の振り直しを行います。")
        mpbNewsSeasonEvent = AddRowText(mpbNewsSeasonEvent, "- - - - - - - - - -")
        mpbNewsSeasonEvent = AddRowText(mpbNewsSeasonEvent, "1位: " & Left(Sheets(season & "_各種記録").Cells(2, "B").Value, 1) & " → 3.5")
        mpbNewsSeasonEvent = AddRowText(mpbNewsSeasonEvent, "2位: " & Left(Sheets(season & "_各種記録").Cells(3, "B").Value, 1) & " → 4.0")
        mpbNewsSeasonEvent = AddRowText(mpbNewsSeasonEvent, "3位: " & Left(Sheets(season & "_各種記録").Cells(4, "B").Value, 1) & " → 4.5")
        mpbNewsSeasonEvent = AddRowText(mpbNewsSeasonEvent, "4位: " & Left(Sheets(season & "_各種記録").Cells(5, "B").Value, 1) & " → 5.0")
        mpbNewsSeasonEvent = AddRowText(mpbNewsSeasonEvent, "5位: " & Left(Sheets(season & "_各種記録").Cells(6, "B").Value, 1) & " → 5.5")
        mpbNewsSeasonEvent = AddRowText(mpbNewsSeasonEvent, "※同率チーム発生時には、必ずしもこの通りとならない場合があります。")
        mpbNewsSeasonEvent = AddRowText(mpbNewsSeasonEvent, "")
        
        tsobBorderDict.Add Left(Sheets(season & "_各種記録").Cells(2, "B").Value, 1), "3.5"
        tsobBorderDict.Add Left(Sheets(season & "_各種記録").Cells(3, "B").Value, 1), "4.0"
        tsobBorderDict.Add Left(Sheets(season & "_各種記録").Cells(4, "B").Value, 1), "4.5"
        tsobBorderDict.Add Left(Sheets(season & "_各種記録").Cells(5, "B").Value, 1), "5.0"
        tsobBorderDict.Add Left(Sheets(season & "_各種記録").Cells(6, "B").Value, 1), "5.5"
        
        Sheets(season & "_スケジュール").Cells(27, "CP").Value = tsobBorderDict.Item(Sheets(season & "_スケジュール").Cells(1, "BB").Value)
        Sheets(season & "_スケジュール").Cells(27, "CQ").Value = tsobBorderDict.Item(Sheets(season & "_スケジュール").Cells(1, "BC").Value)
        Sheets(season & "_スケジュール").Cells(27, "CR").Value = tsobBorderDict.Item(Sheets(season & "_スケジュール").Cells(1, "BD").Value)
        Sheets(season & "_スケジュール").Cells(27, "CS").Value = tsobBorderDict.Item(Sheets(season & "_スケジュール").Cells(1, "BE").Value)
        Sheets(season & "_スケジュール").Cells(27, "CT").Value = tsobBorderDict.Item(Sheets(season & "_スケジュール").Cells(1, "BF").Value)
    End If
    
    ' HDCP変更受付開始
    If section = 10 Or section = 20 Then
        mpbNewsSeasonEventFlg = True
        
        mpbNewsSeasonEvent = AddRowText(mpbNewsSeasonEvent, "・只今より、後半戦からのHDCP変更受付を開始します。第15節終了をもって締め切るので、変更したいチームは、必要に応じて申請を行ってください。変更しない場合は、特に対応不要です。")
        mpbNewsSeasonEvent = AddRowText(mpbNewsSeasonEvent, "")
    End If
    
    ' HDCP変更中
    If section = 11 Or section = 12 Or section = 13 Or section = 14 Then
        mpbNewsSeasonEventFlg = True
        
        mpbNewsSeasonEvent = AddRowText(mpbNewsSeasonEvent, "・後半戦からのHDCP変更を受付中です。変更したいチームは、第15節終了までに申請を行ってください。")
        mpbNewsSeasonEvent = AddRowText(mpbNewsSeasonEvent, "")
    End If
    
    ' HDCP変更受付〆
    If section = 15 Then
        mpbNewsSeasonEventFlg = True
        
        mpbNewsSeasonEvent = AddRowText(mpbNewsSeasonEvent, "・只今をもちまして、後半戦に向けたHDCP変更の申請を締め切ります。HDCPの表示設定を最新化してください。")
        mpbNewsSeasonEvent = AddRowText(mpbNewsSeasonEvent, "")
    End If
    
    ' B9GG提出受付開始
    If section = 25 Then
        mpbNewsSeasonEventFlg = True
        
        mpbNewsSeasonEvent = AddRowText(mpbNewsSeasonEvent, "・只今より、B9GGノミネートオーダーの提出受付を開始します。第28節終了をもって締め切るので、各チーム、LINEグループのアルバム「" & season & "B9GGノミネート」に提出をお願いいたします。")
        mpbNewsSeasonEvent = AddRowText(mpbNewsSeasonEvent, "")
    End If
    
    ' B9GG提出受付中
    If section = 26 Or section = 27 Then
        mpbNewsSeasonEventFlg = True
        
        mpbNewsSeasonEvent = AddRowText(mpbNewsSeasonEvent, "・B9GGノミネートオーダーの提出/変更を受付中です。未提出のチームは、第28節が終了するまでに、LINEグループのアルバム「" & season & "B9GGノミネート」への提出をお願いいたします。")
        mpbNewsSeasonEvent = AddRowText(mpbNewsSeasonEvent, "")
    End If
    
    ' B9GG提出受付〆
    If section = 28 Then
        mpbNewsSeasonEventFlg = True
        
        mpbNewsSeasonEvent = AddRowText(mpbNewsSeasonEvent, "・只今をもちまして、B9GGノミネートオーダーの提出を締め切ります。")
        mpbNewsSeasonEvent = AddRowText(mpbNewsSeasonEvent, "")
    End If
    
    ' MPBアワード案内
    If section = 30 Then
        mpbNewsSeasonEventFlg = True
        
        mpbNewsSeasonEvent = AddRowText(mpbNewsSeasonEvent, "・今シーズン、予定されていた全日程が終了しました。まずは、皆さんお疲れさまでした！この後、MPBアワードを実施しますので、案内をお待ちください。")
        mpbNewsSeasonEvent = AddRowText(mpbNewsSeasonEvent, "")
    End If
    
    ' 結果の出力
    If mpbNewsSeasonEventFlg Then
        mpbNewsSeasonEvent = AddRowText(mpbNewsSeasonEvent, "以上")
        
        If Not debugModeFlg Then
            Call OutputText(mpbNewsSeasonEvent, MPB_WORK_DIRECTORY_PATH & "\mpbnews-seasonevent.txt")
        Else
            Call MessageInfo(mpbNewsSeasonEvent, "MakeMPBNewsSeasonEvent")
        End If
    End If

End Function

' 節の進行により発生する、優勝マジックや自力優勝に関するイベントを出力
Function MakeMPBNewsOfThisSection()
    
    ' 実行条件
    If section = 0 Or True Then
        Exit Function
    End If
    
    ' 宣言
    Dim mpbNewsOfThisSectionFlg As Boolean
    Dim mpbNewsOfThisSection As String
    Dim seasonStatus As New Dictionary
    
    ' 初期化
    mpbNewsOfThisSectionFlg = False
    mpbNewsOfThisSection = "【MPBニュース】"
    
    ' 状況確認(今節実施前)
    seasonStatus.Add "今節実施前", CheckSeasonStatus(section - 1, [["","",""],["","",""]])
    
    ' 今節実施前に優勝が決まっている場合はスキップ
    If seasonStatus.Item("今節実施前")(0) <> "" Then
        Exit Function
    End If
    
    ' 状況確認(今節実施後)
    seasonStatus.Add "今節実施後", CheckSeasonStatus(section, [["","",""],["","",""]])
    
    ' 次節を考える必要がない場合
    If seasonStatus.Item("今節実施後")(0) <> "" Or section = 30 Then
        Dim teamID As Integer
        For teamID = 1 To 5
            If seasonStatus.Item("今節実施後")(teamID) = "優勝" Then
                mpbNewsOfThisSectionFlg = True
                mpbNewsOfThisSection = AddRowText(mpbNewsOfThisSection, "◇" & DICT_TEAMNAME.Item(dictTeamID.Item(teamID)) & "◇MPB(" & season & ")優勝が確定！")
            End If
        Next teamID
    End If
    
    ' 状況確認(次節実施後)
    If Not mpbNewsOfThisSectionFlg Then
        seasonStatus.Add "次節◯-●/◯-●", CheckSeasonStatus(section + 1, [["9","-","0"],["9","-","0"]])
        seasonStatus.Add "次節◯-●/●-◯", CheckSeasonStatus(section + 1, [["9","-","0"],["0","-","9"]])
        seasonStatus.Add "次節●-◯/◯-●", CheckSeasonStatus(section + 1, [["0","-","9"],["9","-","0"]])
        seasonStatus.Add "次節●-◯/●-◯", CheckSeasonStatus(section + 1, [["0","-","9"],["0","-","9"]])
    End If
    
    ' Coming Soon
    
    ' 結果の出力
    If mpbNewsOfThisSectionFlg Then
        If Not debugModeFlg Then
            Call OutputText(mpbNewsOfThisSection, MPB_WORK_DIRECTORY_PATH & "\mpbnews-section.txt")
        Else
            Call MessageInfo(mpbNewsOfThisSection, "MakeMPBNewsOfThisSection")
        End If
    End If
    
End Function

Function CheckSeasonStatus(sectionNumber As Integer, ByRef score As String) As String()
    
    Dim tmp(2) As String
    Dim resultArray(5) As String
    
    If sectionNumber < section Then
        tmp(1) = Sheets(season & "_スケジュール").Cells(sectionNumber * 8 + 3, "F").Value
        tmp(2) = Sheets(season & "_スケジュール").Cells(sectionNumber * 8 + 7, "F").Value
        Sheets(season & "_スケジュール").Cells(sectionNumber * 8 + 3, "F").Value = ""
        Sheets(season & "_スケジュール").Cells(sectionNumber * 8 + 7, "F").Value = ""
    ElseIf sectionNumber > section Then
        Sheets(season & "_スケジュール").Cells(sectionNumber * 8 + 3, "D").Value = score(0, 0)
        Sheets(season & "_スケジュール").Cells(sectionNumber * 8 + 3, "F").Value = score(0, 1)
        Sheets(season & "_スケジュール").Cells(sectionNumber * 8 + 3, "H").Value = score(0, 2)
        Sheets(season & "_スケジュール").Cells(sectionNumber * 8 + 7, "D").Value = score(1, 0)
        Sheets(season & "_スケジュール").Cells(sectionNumber * 8 + 7, "F").Value = score(1, 1)
        Sheets(season & "_スケジュール").Cells(sectionNumber * 8 + 7, "H").Value = score(1, 2)
    End If
    
    Application.Calculate
    
    Dim teamID As Integer
    resultArray(0) = ""
    For teamID = 1 To 5
        
        resultArray(teamID) = "-"
        
        If Sheets(seasonName & "_各種記録").Cells(teamID + 1, "BR").Value = 0 Then
            resultArray(teamID) = "自力V消滅"
        ElseIf Sheets(seasonName & "_各種記録").Cells(teamID + 1, "BX").Value = "優勝" Then
            resultArray(teamID) = Sheets(seasonName & "_各種記録").Cells(teamID + 1, "BX").Value
            resultArray(0) = "優勝チーム決定"
        ElseIf Sheets(seasonName & "_各種記録").Cells(teamID + 1, "BX").Value <> "-" Then
            resultArray(teamID) = Sheets(seasonName & "_各種記録").Cells(teamID + 1, "BX").Value
        End If
        
    Next teamID
    
    If sectionNumber < section Then
        Sheets(season & "_スケジュール").Cells(sectionNumber * 8 + 3, "F").Value = tmp(1)
        Sheets(season & "_スケジュール").Cells(sectionNumber * 8 + 7, "F").Value = tmp(2)
    ElseIf sectionNumber > section Then
        Sheets(season & "_スケジュール").Cells(sectionNumber * 8 + 3, "D").Value = ""
        Sheets(season & "_スケジュール").Cells(sectionNumber * 8 + 3, "F").Value = ""
        Sheets(season & "_スケジュール").Cells(sectionNumber * 8 + 3, "H").Value = ""
        Sheets(season & "_スケジュール").Cells(sectionNumber * 8 + 7, "D").Value = ""
        Sheets(season & "_スケジュール").Cells(sectionNumber * 8 + 7, "F").Value = ""
        Sheets(season & "_スケジュール").Cells(sectionNumber * 8 + 7, "H").Value = ""
    End If
    
    Application.Calculate
    
    CheckSeasonStatus = resultArray()
    
End Function

' スペ判定・結果を出力
Function MakeMPBNewsOfAccident()
    
    ' 実行条件
    If section = 30 Then
        Exit Function
    End If
    
    ' 宣言
    Dim mpbNewsOfAccidentFlg As Boolean
    Dim mpbNewsOfAccident As String
    Dim gamesBeforeThisSection As Integer
    Dim gamesAfterThisSection As Integer
    Dim teamBasedAccidentRate As Single
    
    ' 初期化
    mpbNewsOfAccidentFlg = False
    mpbNewsOfAccident = "【MPBニュース】"
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
    Dim accidentInformationFile As String
    Dim accidentInformationNews As String
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
        Call MessageInfo(dictTeamID.Item(teamID) & " : teamBasedAccidentRate = " & teamBasedAccidentRate * 100 & "%", "MakeMPBNewsOfAccident")
        
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
                dice = 1
            End If
            If dice < teamBasedAccidentRate * DICT_ACCIDENT_COEFFICIENT.Item(Sheets(season & "_投手データ").Cells(rowIdx, "E").Value) Then
                
                ' スペ長さ(表)抽選
                visibleAccidentPeriod = DrawFromDict(DICT_ACCIDENT_LENGTH_RATE)
                
                ' スペ長さ(裏)抽選 ※長さゼロにはならない
                hiddenAccidentPeriod = visibleAccidentPeriod + DrawFromDict(DICT_ACCIDENT_MARGIN_DICT.Item(Sheets(season & "_投手データ").Cells(rowIdx, "E").Value))
                If hiddenAccidentPeriod = 0 Then
                    hiddenAccidentPeriod = 1
                End If
                
                ' スペ内容抽選
                accidentInformation = DrawFromDict(DICT_ACCIDENT_INFORMATION_PITCHER_DICT.Item(visibleAccidentPeriod))
                accidentInformationFile = Split(accidentInformation, "_")(0)
                accidentInformationNews = Split(accidentInformation, "_")(1)
                mpbNewsOfAccident = AddRowText(mpbNewsOfAccident, "◇" & DICT_TEAMNAME.Item(dictTeamID.Item(teamID)) & "◇" & Sheets(season & "_投手データ").Cells(rowIdx, "D").Value & "選手が" & accidentInformationNews)
                mpbNewsOfAccidentFlg = True
                
                ' ファイル書き込み
                For columnIdx = 282 + gamesAfterThisSection To 282 + gamesAfterThisSection + hiddenAccidentPeriod - 1
                    If columnIdx > 305 Then
                        Exit For
                    End If
                    Call MessageDebug(Sheets(season & "_投手データ").Cells(rowIdx, "D").Value & ":" & accidentInformationFile & "(" & visibleAccidentPeriod & ")", "INPUT 投手データ.Cells(" & rowIdx & "," & columnIdx & ")")
                    Sheets(season & "_投手データ").Cells(rowIdx, columnIdx).Value = Sheets(season & "_投手データ").Cells(rowIdx, "D").Value & ":" & accidentInformationFile & "(" & visibleAccidentPeriod & ")"
                Next columnIdx
            
            ElseIf Sheets(season & "_投手データ").Cells(rowIdx, 282 + gamesBeforeThisSection).Value <> "" And Sheets(season & "_投手データ").Cells(rowIdx, 282 + gamesAfterThisSection).Value = "" Then
                
                ' 復帰
                mpbNewsOfAccident = AddRowText(mpbNewsOfAccident, "◇" & DICT_TEAMNAME.Item(dictTeamID.Item(teamID)) & "◇離脱中の" & Sheets(season & "_投手データ").Cells(rowIdx, "D").Value & "選手について、次節からの戦列復帰が明言されました。")
                mpbNewsOfAccidentFlg = True
                
            End If
            
        Next rowIdx
        
        ' 野手スペ判定
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
            If dice < teamBasedAccidentRate * DICT_ACCIDENT_COEFFICIENT.Item(Sheets(season & "_野手データ").Cells(rowIdx, "E").Value) Then
                
                ' スペ長さ(表)抽選
                visibleAccidentPeriod = DrawFromDict(DICT_ACCIDENT_LENGTH_RATE)
                
                ' スペ長さ(裏)抽選 ※長さゼロにはならない
                hiddenAccidentPeriod = visibleAccidentPeriod + DrawFromDict(DICT_ACCIDENT_MARGIN_DICT.Item(Sheets(season & "_野手データ").Cells(rowIdx, "E").Value))
                If hiddenAccidentPeriod = 0 Then
                    hiddenAccidentPeriod = 1
                End If
                
                ' スペ内容抽選
                accidentInformation = DrawFromDict(DICT_ACCIDENT_INFORMATION_FIELDER_DICT.Item(visibleAccidentPeriod))
                accidentInformationFile = Split(accidentInformation, "_")(0)
                accidentInformationNews = Split(accidentInformation, "_")(1)
                mpbNewsOfAccident = AddRowText(mpbNewsOfAccident, "◇" & DICT_TEAMNAME.Item(dictTeamID.Item(teamID)) & "◇" & Sheets(season & "_野手データ").Cells(rowIdx, "D").Value & "選手が" & accidentInformationNews)
                mpbNewsOfAccidentFlg = True
                
                ' ファイル書き込み
                For columnIdx = 236 + gamesAfterThisSection To 236 + gamesAfterThisSection + hiddenAccidentPeriod - 1
                    If columnIdx > 259 Then
                        Exit For
                    End If
                    Call MessageDebug(Sheets(season & "_野手データ").Cells(rowIdx, "D").Value & ":" & accidentInformationFile & "(" & visibleAccidentPeriod & ")", "INPUT 野手データ.Cells(" & rowIdx & "," & columnIdx & ")")
                    Sheets(season & "_野手データ").Cells(rowIdx, columnIdx).Value = Sheets(season & "_野手データ").Cells(rowIdx, "D").Value & ":" & accidentInformationFile & "(" & visibleAccidentPeriod & ")"
                Next columnIdx
                
                
            ElseIf Sheets(season & "_野手データ").Cells(rowIdx, 236 + gamesBeforeThisSection).Value <> "" And Sheets(season & "_野手データ").Cells(rowIdx, 236 + gamesAfterThisSection).Value = "" Then

                ' 復帰
                mpbNewsOfAccident = AddRowText(mpbNewsOfAccident, "◇" & DICT_TEAMNAME.Item(dictTeamID.Item(teamID)) & "◇離脱中の" & Sheets(season & "_野手データ").Cells(rowIdx, "D").Value & "選手について、次節からの戦列復帰が明言されました。")
                mpbNewsOfAccidentFlg = True
                
            End If
        
        Next rowIdx
        
    Next teamID
    
    ' 結果の出力
    If mpbNewsOfAccidentFlg Then
        If Not debugModeFlg Then
            Call OutputText(mpbNewsOfAccident, MPB_WORK_DIRECTORY_PATH & "\mpbnews-accident.txt")
        Else
            Call MessageInfo(mpbNewsOfAccident, "MakeMPBNewsOfAccident")
        End If
    End If
    
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

