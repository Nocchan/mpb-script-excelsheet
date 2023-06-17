Attribute VB_Name = "mpb_vbascript_sectionCompletion"
Sub アクシデント発生()

    Dim debugModeFlg As Boolean
    debugModeFlg = False
    If debugModeFlg Then
        MsgBox "デバッグモード"
    End If
    
    ' エラーチェック
    If ActiveSheet.Name <> ActiveSheet.Cells(1, "A").Value & "_スケジュール" Then
        MsgBox "シート名またはA1セルのシーズン指定が不正です。"
        End
    End If
    
    If Not debugModeFlg Then
        Application.ScreenUpdating = False
    End If
    
    ' 全体で使用する変数
    Dim seasonName As String
    Dim numberOfSection As Integer
    Dim allAccidentResultNotification, tsobChangeNotification, rankNotification As String
    
    ' チームごとに使用する変数
    Dim teamID As Integer
    Dim teamName As String
    Dim numberOfGamesPlayedBeforeThisSection, numberOfGamesPlayedAfterThisSection As Integer
    Dim pcrPositiveRate, accidentBonusValue As Single
    Dim pcrResultText, pcrResultMessage As String
    
    ' 選手ごとに使用する変数
    Dim playerID, rowIdx As Integer
    Dim playerName, playerNameRegistered As String
    Dim accidentRate As Single
    Dim accidentPeriod As Integer
    Dim accidentPeriodString, accidentOverview, accidentText, accidentMessage As String
    
    ' その他変数
    Dim columnIdxOfStamina, columnIdx As Integer
    Dim dice As Single
    
    Randomize
    
    ' シーズン、節進行状況の確認
    seasonName = ActiveSheet.Cells(1, "A").Value
    numberOfSection = WorksheetFunction.CountIf(ActiveSheet.Range("BA2:BA241"), 0) / 8
    If numberOfSection = 0 Then
        allAccidentResultNotification = "【開幕前スペ判定】"
    Else
        allAccidentResultNotification = "【第" & numberOfSection & "節終了時スペ判定】"
    End If
    rankNotification = "【MPBニュース】"
    If numberOfSection = 25 Then
        rankNotification = rankNotification & vbCrLf & _
                           "・<重要>只今より、B9GGノミネートオーダーの提出受付を開始します。〆切は第28節終了時です。各チーム、LINEグループのアルバムに提出をお願いいたします。"
    ElseIf numberOfSection = 26 Or numberOfSection = 27 Then
        rankNotification = rankNotification & vbCrLf & _
                           "・<重要>B9GGノミネートオーダーを提出受付中です。未提出のチームは、第28節が終了するまでに、LINEグループのアルバムへの提出をお願いいたします。"
    ElseIf numberOfSection = 28 Then
        rankNotification = rankNotification & vbCrLf & _
                           "・<重要>B9GGノミネートオーダーの提出/変更受付を締め切りました。"
    End If
    tsobChangeNotification = "【TS/OB枠振り直しのお知らせ】"
    
    If Not debugModeFlg Then
        ' 次節の未開始と予告先発の出揃いを確認
        If numberOfSection > 0 Then
            If ActiveSheet.Cells(numberOfSection * 8 + 2, "D").Value = "" Or ActiveSheet.Cells(numberOfSection * 8 + 6, "D").Value = "" Or _
               ActiveSheet.Cells(numberOfSection * 8 + 2, "H").Value = "" Or ActiveSheet.Cells(numberOfSection * 8 + 6, "H").Value = "" Then
                MsgBox "第" & numberOfSection + 1 & "節の先発予告が完了していません。"
                End
            End If
        End If
        
        If ActiveSheet.Cells(numberOfSection * 8 + 3, "D").Value <> "" Or ActiveSheet.Cells(numberOfSection * 8 + 7, "D").Value <> "" Or _
           ActiveSheet.Cells(numberOfSection * 8 + 3, "F").Value <> "" Or ActiveSheet.Cells(numberOfSection * 8 + 7, "F").Value <> "" Or _
           ActiveSheet.Cells(numberOfSection * 8 + 3, "H").Value <> "" Or ActiveSheet.Cells(numberOfSection * 8 + 7, "H").Value <> "" Then
            MsgBox "第" & numberOfSection + 1 & "節の試合結果が不正に入力されています。"
            End
        End If
    End If
    
    ' ここに追記
    ' 例外的にここで変数宣言を行う
    Dim tmp1, tmp2 As String
    Dim vStatus(6, 6) As String ' [0:今節実施前,1:今節実施後,2:次節◯-●/◯-●,3:次節◯-●/●-◯,4:次節●-◯/◯-●5:次節●-◯/●-◯][teamID+flag]∈{優勝,M*,自力V消滅,-}
    Dim teamNameOfNextSection(2, 2) As String ' [0:①,1:②][0:Home,1:Visitor]
    
    If numberOfSection > 0 And numberOfSection < 30 Then
        ' 今節実施前の状況確認
        tmp1 = ActiveSheet.Cells(numberOfSection * 8 - 5, "F").Value
        tmp2 = ActiveSheet.Cells(numberOfSection * 8 - 1, "F").Value
        ActiveSheet.Cells(numberOfSection * 8 - 5, "F").Value = ""
        ActiveSheet.Cells(numberOfSection * 8 - 1, "F").Value = ""
        Application.Calculate
        
        vStatus(0, 5) = "false"
        
        For teamID = 0 To 4
            
            vStatus(0, teamID) = "-"
            
            If Sheets(seasonName & "_各種記録").Cells(teamID + 2, "BR").Value = 0 Then
                vStatus(0, teamID) = "自力V消滅"
            ElseIf Sheets(seasonName & "_各種記録").Cells(teamID + 2, "BX").Value <> "-" Then
                vStatus(0, teamID) = Sheets(seasonName & "_各種記録").Cells(teamID + 2, "BX").Value
            End If
            
        Next teamID
        
        ' 今節実施後の状況確認
        ActiveSheet.Cells(numberOfSection * 8 - 5, "F").Value = tmp1
        ActiveSheet.Cells(numberOfSection * 8 - 1, "F").Value = tmp2
        Application.Calculate
        
        vStatus(1, 5) = "false"
        
        For teamID = 0 To 4
            
            vStatus(1, teamID) = "-"
            
            If Sheets(seasonName & "_各種記録").Cells(teamID + 2, "BR").Value = 0 Then
                vStatus(1, teamID) = "自力V消滅"
            ElseIf Sheets(seasonName & "_各種記録").Cells(teamID + 2, "BX").Value <> "-" Then
                vStatus(1, teamID) = Sheets(seasonName & "_各種記録").Cells(teamID + 2, "BX").Value
            End If
            
            If Left(vStatus(0, teamID), 1) <> Left(vStatus(1, teamID), 1) Then
                vStatus(1, 5) = "true"
            End If
            
        Next teamID
        
        ' 次節◯-●/◯-●の状況確認
        ActiveSheet.Cells(numberOfSection * 8 + 3, "D").Value = "X"
        ActiveSheet.Cells(numberOfSection * 8 + 3, "F").Value = "-"
        ActiveSheet.Cells(numberOfSection * 8 + 3, "H").Value = "0"
        ActiveSheet.Cells(numberOfSection * 8 + 7, "D").Value = "X"
        ActiveSheet.Cells(numberOfSection * 8 + 7, "F").Value = "-"
        ActiveSheet.Cells(numberOfSection * 8 + 7, "H").Value = "0"
        Application.Calculate
        
        vStatus(2, 5) = "false"
        
        For teamID = 0 To 4
            
            vStatus(2, teamID) = "-"
            
            If Sheets(seasonName & "_各種記録").Cells(teamID + 2, "BR").Value = 0 Then
                vStatus(2, teamID) = "自力V消滅"
            ElseIf Sheets(seasonName & "_各種記録").Cells(teamID + 2, "BX").Value <> "-" Then
                vStatus(2, teamID) = Sheets(seasonName & "_各種記録").Cells(teamID + 2, "BX").Value
            End If
            
            If Left(vStatus(1, teamID), 1) <> Left(vStatus(2, teamID), 1) Then
                vStatus(2, 5) = "true"
            End If
            
        Next teamID
        
        ' 次節◯-●/●-◯の状況確認
        ActiveSheet.Cells(numberOfSection * 8 + 3, "D").Value = "X"
        ActiveSheet.Cells(numberOfSection * 8 + 3, "F").Value = "-"
        ActiveSheet.Cells(numberOfSection * 8 + 3, "H").Value = "0"
        ActiveSheet.Cells(numberOfSection * 8 + 7, "D").Value = "0"
        ActiveSheet.Cells(numberOfSection * 8 + 7, "F").Value = "-"
        ActiveSheet.Cells(numberOfSection * 8 + 7, "H").Value = "X"
        Application.Calculate
        
        vStatus(3, 5) = "false"
        
        For teamID = 0 To 4
            
            vStatus(3, teamID) = "-"
            
            If Sheets(seasonName & "_各種記録").Cells(teamID + 2, "BR").Value = 0 Then
                vStatus(3, teamID) = "自力V消滅"
            ElseIf Sheets(seasonName & "_各種記録").Cells(teamID + 2, "BX").Value <> "-" Then
                vStatus(3, teamID) = Sheets(seasonName & "_各種記録").Cells(teamID + 2, "BX").Value
            End If
            
            If Left(vStatus(1, teamID), 1) <> Left(vStatus(3, teamID), 1) Then
                vStatus(3, 5) = "true"
            End If
            
        Next teamID
        
        ' 次節●-◯/◯-●の状況確認
        ActiveSheet.Cells(numberOfSection * 8 + 3, "D").Value = "0"
        ActiveSheet.Cells(numberOfSection * 8 + 3, "F").Value = "-"
        ActiveSheet.Cells(numberOfSection * 8 + 3, "H").Value = "X"
        ActiveSheet.Cells(numberOfSection * 8 + 7, "D").Value = "X"
        ActiveSheet.Cells(numberOfSection * 8 + 7, "F").Value = "-"
        ActiveSheet.Cells(numberOfSection * 8 + 7, "H").Value = "0"
        Application.Calculate
        
        vStatus(4, 5) = "false"
        
        For teamID = 0 To 4
            
            vStatus(4, teamID) = "-"
            
            If Sheets(seasonName & "_各種記録").Cells(teamID + 2, "BR").Value = 0 Then
                vStatus(4, teamID) = "自力V消滅"
            ElseIf Sheets(seasonName & "_各種記録").Cells(teamID + 2, "BX").Value <> "-" Then
                vStatus(4, teamID) = Sheets(seasonName & "_各種記録").Cells(teamID + 2, "BX").Value
            End If
            
            If Left(vStatus(1, teamID), 1) <> Left(vStatus(4, teamID), 1) Then
                vStatus(4, 5) = "true"
            End If
            
        Next teamID
        
        ' 次節●-◯/●-◯の状況確認
        ActiveSheet.Cells(numberOfSection * 8 + 3, "D").Value = "0"
        ActiveSheet.Cells(numberOfSection * 8 + 3, "F").Value = "-"
        ActiveSheet.Cells(numberOfSection * 8 + 3, "H").Value = "X"
        ActiveSheet.Cells(numberOfSection * 8 + 7, "D").Value = "0"
        ActiveSheet.Cells(numberOfSection * 8 + 7, "F").Value = "-"
        ActiveSheet.Cells(numberOfSection * 8 + 7, "H").Value = "X"
        Application.Calculate
        
        vStatus(5, 5) = "false"
        
        For teamID = 0 To 4
            
            vStatus(5, teamID) = "-"
            
            If Sheets(seasonName & "_各種記録").Cells(teamID + 2, "BR").Value = 0 Then
                vStatus(5, teamID) = "自力V消滅"
            ElseIf Sheets(seasonName & "_各種記録").Cells(teamID + 2, "BX").Value <> "-" Then
                vStatus(5, teamID) = Sheets(seasonName & "_各種記録").Cells(teamID + 2, "BX").Value
            End If
            
            If Left(vStatus(1, teamID), 1) <> Left(vStatus(5, teamID), 1) Then
                vStatus(5, 5) = "true"
            End If
            
        Next teamID
        
        ActiveSheet.Cells(numberOfSection * 8 + 3, "D").Value = ""
        ActiveSheet.Cells(numberOfSection * 8 + 3, "F").Value = ""
        ActiveSheet.Cells(numberOfSection * 8 + 3, "H").Value = ""
        ActiveSheet.Cells(numberOfSection * 8 + 7, "D").Value = ""
        ActiveSheet.Cells(numberOfSection * 8 + 7, "F").Value = ""
        ActiveSheet.Cells(numberOfSection * 8 + 7, "H").Value = ""
        Application.Calculate
        
        ' 今節までに優勝が決まっている場合スキップ
        For teamID = 0 To 4
            
            If vStatus(0, teamID) = "優勝" Then
                GoTo MPB_NEWS_CHECK_END_POINT
            ElseIf vStatus(1, teamID) = "優勝" Then
                rankNotification = rankNotification & vbCrLf & _
                                   "・" & Sheets(seasonName & "_各種記録").Cells(teamID + 2, "BA").Value & ":リーグ優勝が確定！"
                GoTo MPB_NEWS_CHECK_END_POINT
            End If
            
        Next teamID
        
        If vStatus(1, 5) = "true" Then
            ' 今節のマジック消滅
            For teamID = 0 To 4
                
                If Left(vStatus(0, teamID), 1) = "M" And Left(vStatus(1, teamID), 1) <> "M" Then
                    rankNotification = rankNotification & vbCrLf & _
                                       "・" & Sheets(seasonName & "_各種記録").Cells(teamID + 2, "BA").Value & ":優勝マジックが消滅…"
                End If
                
            Next teamID
            
            ' 今節の自力V消滅
            For teamID = 0 To 4
                
                If vStatus(0, teamID) <> "自力V消滅" And vStatus(1, teamID) = "自力V消滅" Then
                    rankNotification = rankNotification & vbCrLf & _
                                       "・" & Sheets(seasonName & "_各種記録").Cells(teamID + 2, "BA").Value & ":自力優勝が消滅…"
                End If
                
            Next teamID
            
            ' 今節の自力V復活
            For teamID = 0 To 4
                
                If vStatus(0, teamID) = "自力V消滅" And vStatus(1, teamID) <> "自力V消滅" Then
                    rankNotification = rankNotification & vbCrLf & _
                                       "・" & Sheets(seasonName & "_各種記録").Cells(teamID + 2, "BA").Value & ":自力優勝が復活！"
                End If
                
            Next teamID
            
            ' 今節のマジック点灯
            For teamID = 0 To 4
                
                If Left(vStatus(0, teamID), 1) <> "M" And Left(vStatus(1, teamID), 1) = "M" Then
                    rankNotification = rankNotification & vbCrLf & _
                                       "・" & Sheets(seasonName & "_各種記録").Cells(teamID + 2, "BA").Value & ":優勝マジック(" & vStatus(1, teamID) & ")が点灯！"
                End If
                
            Next teamID
            
        End If
        
        ' 次節実施後の展望
        teamNameOfNextSection(0, 0) = ActiveSheet.Cells(numberOfSection * 8 + 2, "C").Value
        teamNameOfNextSection(0, 1) = ActiveSheet.Cells(numberOfSection * 8 + 2, "J").Value
        teamNameOfNextSection(1, 0) = ActiveSheet.Cells(numberOfSection * 8 + 6, "C").Value
        teamNameOfNextSection(1, 1) = ActiveSheet.Cells(numberOfSection * 8 + 6, "J").Value
        
        ' ①◯-●で共通
        If vStatus(2, 5) = "true" And vStatus(3, 5) = "true" Then
            ' 優勝
            For teamID = 0 To 4
                
                If Left(vStatus(1, teamID), 1) <> "優" And Left(vStatus(2, teamID), 1) = "優" And Left(vStatus(3, teamID), 1) = "優" Then
                    rankNotification = rankNotification & vbCrLf & _
                                       "・次節 " & teamNameOfNextSection(0, 0) & "◯-●" & teamNameOfNextSection(0, 1) & " で、" & vbCrLf & _
                                       "　" & Sheets(seasonName & "_各種記録").Cells(teamID + 2, "BA").Value & "のリーグ優勝が確定！"
                    GoTo MPB_NEWS_PATTERN1_END_POINT
                End If
                
            Next teamID
            
            ' マジック消滅
            For teamID = 0 To 4
                
                If Left(vStatus(1, teamID), 1) = "M" And Left(vStatus(2, teamID), 1) <> "M" And Left(vStatus(3, teamID), 1) <> "M" Then
                    rankNotification = rankNotification & vbCrLf & _
                                       "・次節 " & teamNameOfNextSection(0, 0) & "◯-●" & teamNameOfNextSection(0, 1) & " で、" & vbCrLf & _
                                       "　" & Sheets(seasonName & "_各種記録").Cells(teamID + 2, "BA").Value & "の優勝マジックが消滅…"
                End If
                
            Next teamID
            
            ' 自力V消滅
            For teamID = 0 To 4
                
                If Left(vStatus(1, teamID), 1) <> "自" And Left(vStatus(2, teamID), 1) = "自" And Left(vStatus(3, teamID), 1) = "自" Then
                    rankNotification = rankNotification & vbCrLf & _
                                       "・次節 " & teamNameOfNextSection(0, 0) & "◯-●" & teamNameOfNextSection(0, 1) & " で、" & vbCrLf & _
                                       "　" & Sheets(seasonName & "_各種記録").Cells(teamID + 2, "BA").Value & "の自力優勝が消滅…"
                End If
                
            Next teamID
            
            ' 自力V復活
            For teamID = 0 To 4
                
                If Left(vStatus(1, teamID), 1) = "自" And Left(vStatus(2, teamID), 1) <> "自" And Left(vStatus(3, teamID), 1) <> "自" Then
                    rankNotification = rankNotification & vbCrLf & _
                                       "・次節 " & teamNameOfNextSection(0, 0) & "◯-●" & teamNameOfNextSection(0, 1) & " で、" & vbCrLf & _
                                       "　" & Sheets(seasonName & "_各種記録").Cells(teamID + 2, "BA").Value & "の自力優勝が復活！"
                End If
                
            Next teamID
            
            ' マジック点灯
            For teamID = 0 To 4
                
                If Left(vStatus(1, teamID), 1) <> "M" And Left(vStatus(2, teamID), 1) = "M" And Left(vStatus(3, teamID), 1) = "M" Then
                    rankNotification = rankNotification & vbCrLf & _
                                       "・次節 " & teamNameOfNextSection(0, 0) & "◯-●" & teamNameOfNextSection(0, 1) & " で、" & vbCrLf & _
                                       "　" & Sheets(seasonName & "_各種記録").Cells(teamID + 2, "BA").Value & "の優勝マジックが点灯！"
                End If
                
            Next teamID
            
        End If
        
MPB_NEWS_PATTERN1_END_POINT:

        ' ①●-◯で共通
        If vStatus(4, 5) = "true" And vStatus(5, 5) = "true" Then
            ' 優勝
            For teamID = 0 To 4
                
                If Left(vStatus(1, teamID), 1) <> "優" And Left(vStatus(4, teamID), 1) = "優" And Left(vStatus(5, teamID), 1) = "優" Then
                    rankNotification = rankNotification & vbCrLf & _
                                       "・次節 " & teamNameOfNextSection(0, 0) & "●-◯" & teamNameOfNextSection(0, 1) & " で、" & vbCrLf & _
                                       "　" & Sheets(seasonName & "_各種記録").Cells(teamID + 2, "BA").Value & "のリーグ優勝が確定！"
                    GoTo MPB_NEWS_PATTERN2_END_POINT
                End If
                
            Next teamID
            
            ' マジック消滅
            For teamID = 0 To 4
                
                If Left(vStatus(1, teamID), 1) = "M" And Left(vStatus(4, teamID), 1) <> "M" And Left(vStatus(5, teamID), 1) <> "M" Then
                    rankNotification = rankNotification & vbCrLf & _
                                       "・次節 " & teamNameOfNextSection(0, 0) & "●-◯" & teamNameOfNextSection(0, 1) & " で、" & vbCrLf & _
                                       "　" & Sheets(seasonName & "_各種記録").Cells(teamID + 2, "BA").Value & "の優勝マジックが消滅…"
                End If
                
            Next teamID
            
            ' 自力V消滅
            For teamID = 0 To 4
                
                If Left(vStatus(1, teamID), 1) <> "自" And Left(vStatus(4, teamID), 1) = "自" And Left(vStatus(5, teamID), 1) = "自" Then
                    rankNotification = rankNotification & vbCrLf & _
                                       "・次節 " & teamNameOfNextSection(0, 0) & "●-◯" & teamNameOfNextSection(0, 1) & " で、" & vbCrLf & _
                                       "　" & Sheets(seasonName & "_各種記録").Cells(teamID + 2, "BA").Value & "の自力優勝が消滅…"
                End If
                
            Next teamID
            
            ' 自力V復活
            For teamID = 0 To 4
                
                If Left(vStatus(1, teamID), 1) = "自" And Left(vStatus(4, teamID), 1) <> "自" And Left(vStatus(5, teamID), 1) <> "自" Then
                    rankNotification = rankNotification & vbCrLf & _
                                       "・次節 " & teamNameOfNextSection(0, 0) & "●-◯" & teamNameOfNextSection(0, 1) & " で、" & vbCrLf & _
                                       "　" & Sheets(seasonName & "_各種記録").Cells(teamID + 2, "BA").Value & "の自力優勝が復活！"
                End If
                
            Next teamID
            
            ' マジック点灯
            For teamID = 0 To 4
                
                If Left(vStatus(1, teamID), 1) <> "M" And Left(vStatus(4, teamID), 1) = "M" And Left(vStatus(5, teamID), 1) = "M" Then
                    rankNotification = rankNotification & vbCrLf & _
                                       "・次節 " & teamNameOfNextSection(0, 0) & "●-◯" & teamNameOfNextSection(0, 1) & " で、" & vbCrLf & _
                                       "　" & Sheets(seasonName & "_各種記録").Cells(teamID + 2, "BA").Value & "の優勝マジックが点灯！"
                End If
                
            Next teamID
            
        End If
        
MPB_NEWS_PATTERN2_END_POINT:

        ' ②◯-●で共通
        If vStatus(2, 5) = "true" And vStatus(4, 5) = "true" Then
            ' 優勝
            For teamID = 0 To 4
                
                If Left(vStatus(1, teamID), 1) <> "優" And Left(vStatus(2, teamID), 1) = "優" And Left(vStatus(4, teamID), 1) = "優" Then
                    rankNotification = rankNotification & vbCrLf & _
                                       "・次節 " & teamNameOfNextSection(1, 0) & "◯-●" & teamNameOfNextSection(1, 1) & " で、" & vbCrLf & _
                                       "　" & Sheets(seasonName & "_各種記録").Cells(teamID + 2, "BA").Value & "のリーグ優勝が確定！"
                    GoTo MPB_NEWS_PATTERN3_END_POINT
                End If
                
            Next teamID
            
            ' マジック消滅
            For teamID = 0 To 4
                
                If Left(vStatus(1, teamID), 1) = "M" And Left(vStatus(2, teamID), 1) <> "M" And Left(vStatus(4, teamID), 1) <> "M" Then
                    rankNotification = rankNotification & vbCrLf & _
                                       "・次節 " & teamNameOfNextSection(1, 0) & "◯-●" & teamNameOfNextSection(1, 1) & " で、" & vbCrLf & _
                                       "　" & Sheets(seasonName & "_各種記録").Cells(teamID + 2, "BA").Value & "の優勝マジックが消滅…"
                End If
                
            Next teamID
            
            ' 自力V消滅
            For teamID = 0 To 4
                
                If Left(vStatus(1, teamID), 1) <> "自" And Left(vStatus(2, teamID), 1) = "自" And Left(vStatus(4, teamID), 1) = "自" Then
                    rankNotification = rankNotification & vbCrLf & _
                                       "・次節 " & teamNameOfNextSection(1, 0) & "◯-●" & teamNameOfNextSection(1, 1) & " で、" & vbCrLf & _
                                       "　" & Sheets(seasonName & "_各種記録").Cells(teamID + 2, "BA").Value & "の自力優勝が消滅…"
                End If
                
            Next teamID
            
            ' 自力V復活
            For teamID = 0 To 4
                
                If Left(vStatus(1, teamID), 1) = "自" And Left(vStatus(2, teamID), 1) <> "自" And Left(vStatus(4, teamID), 1) <> "自" Then
                    rankNotification = rankNotification & vbCrLf & _
                                       "・次節 " & teamNameOfNextSection(1, 0) & "◯-●" & teamNameOfNextSection(1, 1) & " で、" & vbCrLf & _
                                       "　" & Sheets(seasonName & "_各種記録").Cells(teamID + 2, "BA").Value & "の自力優勝が復活！"
                End If
                
            Next teamID
            
            ' マジック点灯
            For teamID = 0 To 4
                
                If Left(vStatus(1, teamID), 1) <> "M" And Left(vStatus(2, teamID), 1) = "M" And Left(vStatus(4, teamID), 1) = "M" Then
                    rankNotification = rankNotification & vbCrLf & _
                                       "・次節 " & teamNameOfNextSection(1, 0) & "◯-●" & teamNameOfNextSection(1, 1) & " で、" & vbCrLf & _
                                       "　" & Sheets(seasonName & "_各種記録").Cells(teamID + 2, "BA").Value & "の優勝マジックが点灯！"
                End If
                
            Next teamID
            
        End If
        
MPB_NEWS_PATTERN3_END_POINT:

        ' ②●-◯で共通
        If vStatus(3, 5) = "true" And vStatus(5, 5) = "true" Then
            ' 優勝
            For teamID = 0 To 4
                
                If Left(vStatus(1, teamID), 1) <> "優" And Left(vStatus(3, teamID), 1) = "優" And Left(vStatus(5, teamID), 1) = "優" Then
                    rankNotification = rankNotification & vbCrLf & _
                                       "・次節 " & teamNameOfNextSection(1, 0) & "●-◯" & teamNameOfNextSection(1, 1) & " で、" & vbCrLf & _
                                       "　" & Sheets(seasonName & "_各種記録").Cells(teamID + 2, "BA").Value & "のリーグ優勝が確定！"
                    GoTo MPB_NEWS_PATTERN4_END_POINT
                End If
                
            Next teamID
            
            ' マジック消滅
            For teamID = 0 To 4
                
                If Left(vStatus(1, teamID), 1) = "M" And Left(vStatus(3, teamID), 1) <> "M" And Left(vStatus(5, teamID), 1) <> "M" Then
                    rankNotification = rankNotification & vbCrLf & _
                                       "・次節 " & teamNameOfNextSection(1, 0) & "●-◯" & teamNameOfNextSection(1, 1) & " で、" & vbCrLf & _
                                       "　" & Sheets(seasonName & "_各種記録").Cells(teamID + 2, "BA").Value & "の優勝マジックが消滅…"
                End If
                
            Next teamID
            
            ' 自力V消滅
            For teamID = 0 To 4
                
                If Left(vStatus(1, teamID), 1) <> "自" And Left(vStatus(3, teamID), 1) = "自" And Left(vStatus(5, teamID), 1) = "自" Then
                    rankNotification = rankNotification & vbCrLf & _
                                       "・次節 " & teamNameOfNextSection(1, 0) & "●-◯" & teamNameOfNextSection(1, 1) & " で、" & vbCrLf & _
                                       "　" & Sheets(seasonName & "_各種記録").Cells(teamID + 2, "BA").Value & "の自力優勝が消滅…"
                End If
                
            Next teamID
            
            ' 自力V復活
            For teamID = 0 To 4
                
                If Left(vStatus(1, teamID), 1) = "自" And Left(vStatus(3, teamID), 1) <> "自" And Left(vStatus(5, teamID), 1) <> "自" Then
                    rankNotification = rankNotification & vbCrLf & _
                                       "・次節 " & teamNameOfNextSection(1, 0) & "●-◯" & teamNameOfNextSection(1, 1) & " で、" & vbCrLf & _
                                       "　" & Sheets(seasonName & "_各種記録").Cells(teamID + 2, "BA").Value & "の自力優勝が復活！"
                End If
                
            Next teamID
            
            ' マジック点灯
            For teamID = 0 To 4
                
                If Left(vStatus(1, teamID), 1) <> "M" And Left(vStatus(3, teamID), 1) = "M" And Left(vStatus(5, teamID), 1) = "M" Then
                    rankNotification = rankNotification & vbCrLf & _
                                       "・次節 " & teamNameOfNextSection(1, 0) & "●-◯" & teamNameOfNextSection(1, 1) & " で、" & vbCrLf & _
                                       "　" & Sheets(seasonName & "_各種記録").Cells(teamID + 2, "BA").Value & "の優勝マジックが点灯！"
                End If
                
            Next teamID
            
        End If
        
MPB_NEWS_PATTERN4_END_POINT:

        ' ①◯-●②◯-●
        If vStatus(2, 5) = "true" Then
            ' 優勝
            For teamID = 0 To 4
                
                If Left(vStatus(1, teamID), 1) <> "優" And Left(vStatus(2, teamID), 1) = "優" And Left(vStatus(3, teamID), 1) <> "優" And Left(vStatus(4, teamID), 1) <> "優" Then
                    rankNotification = rankNotification & vbCrLf & _
                                       "・次節 " & teamNameOfNextSection(0, 0) & "◯-●" & teamNameOfNextSection(0, 1) & " & " & teamNameOfNextSection(1, 0) & "◯-●" & teamNameOfNextSection(1, 1) & " で、" & vbCrLf & _
                                       "　" & Sheets(seasonName & "_各種記録").Cells(teamID + 2, "BA").Value & "のリーグ優勝が確定！"
                    GoTo MPB_NEWS_PATTERN5_END_POINT
                End If
                
            Next teamID
            
            ' マジック消滅
            For teamID = 0 To 4
                
                If Left(vStatus(1, teamID), 1) = "M" And Left(vStatus(2, teamID), 1) <> "M" And Left(vStatus(3, teamID), 1) = "M" And Left(vStatus(4, teamID), 1) = "M" Then
                    rankNotification = rankNotification & vbCrLf & _
                                       "・次節 " & teamNameOfNextSection(0, 0) & "◯-●" & teamNameOfNextSection(0, 1) & " & " & teamNameOfNextSection(1, 0) & "◯-●" & teamNameOfNextSection(1, 1) & " で、" & vbCrLf & _
                                       "　" & Sheets(seasonName & "_各種記録").Cells(teamID + 2, "BA").Value & "の優勝マジックが消滅…"
                End If
                
            Next teamID
            
            ' 自力V消滅
            For teamID = 0 To 4
                
                If Left(vStatus(1, teamID), 1) <> "自" And Left(vStatus(2, teamID), 1) = "自" And Left(vStatus(3, teamID), 1) <> "自" And Left(vStatus(4, teamID), 1) <> "自" Then
                    rankNotification = rankNotification & vbCrLf & _
                                       "・次節 " & teamNameOfNextSection(0, 0) & "◯-●" & teamNameOfNextSection(0, 1) & " & " & teamNameOfNextSection(1, 0) & "◯-●" & teamNameOfNextSection(1, 1) & " で、" & vbCrLf & _
                                       "　" & Sheets(seasonName & "_各種記録").Cells(teamID + 2, "BA").Value & "の自力優勝が消滅…"
                End If
                
            Next teamID
            
            ' 自力V復活
            For teamID = 0 To 4
                
                If Left(vStatus(1, teamID), 1) = "自" And Left(vStatus(2, teamID), 1) <> "自" And Left(vStatus(3, teamID), 1) = "自" And Left(vStatus(4, teamID), 1) = "自" Then
                    rankNotification = rankNotification & vbCrLf & _
                                       "・次節 " & teamNameOfNextSection(0, 0) & "◯-●" & teamNameOfNextSection(0, 1) & " & " & teamNameOfNextSection(1, 0) & "◯-●" & teamNameOfNextSection(1, 1) & " で、" & vbCrLf & _
                                       "　" & Sheets(seasonName & "_各種記録").Cells(teamID + 2, "BA").Value & "の自力優勝が復活！"
                End If
                
            Next teamID
            
            ' マジック点灯
            For teamID = 0 To 4
                
                If Left(vStatus(1, teamID), 1) <> "M" And Left(vStatus(2, teamID), 1) = "M" And Left(vStatus(3, teamID), 1) <> "M" And Left(vStatus(4, teamID), 1) <> "M" Then
                    rankNotification = rankNotification & vbCrLf & _
                                       "・次節 " & teamNameOfNextSection(0, 0) & "◯-●" & teamNameOfNextSection(0, 1) & " & " & teamNameOfNextSection(1, 0) & "◯-●" & teamNameOfNextSection(1, 1) & " で、" & vbCrLf & _
                                       "　" & Sheets(seasonName & "_各種記録").Cells(teamID + 2, "BA").Value & "の優勝マジックが点灯！"
                End If
                
            Next teamID
            
        End If
        
MPB_NEWS_PATTERN5_END_POINT:

        ' ①◯-●②●-◯
        If vStatus(3, 5) = "true" Then
            ' 優勝
            For teamID = 0 To 4
                
                If Left(vStatus(1, teamID), 1) <> "優" And Left(vStatus(2, teamID), 1) <> "優" And Left(vStatus(3, teamID), 1) = "優" And Left(vStatus(5, teamID), 1) <> "優" Then
                    rankNotification = rankNotification & vbCrLf & _
                                       "・次節 " & teamNameOfNextSection(0, 0) & "◯-●" & teamNameOfNextSection(0, 1) & " & " & teamNameOfNextSection(1, 0) & "●-◯" & teamNameOfNextSection(1, 1) & " で、" & vbCrLf & _
                                       "　" & Sheets(seasonName & "_各種記録").Cells(teamID + 2, "BA").Value & "のリーグ優勝が確定！"
                    GoTo MPB_NEWS_PATTERN6_END_POINT
                End If
                
            Next teamID
            
            ' マジック消滅
            For teamID = 0 To 4
                
                If Left(vStatus(1, teamID), 1) = "M" And Left(vStatus(2, teamID), 1) = "M" And Left(vStatus(3, teamID), 1) <> "M" And Left(vStatus(5, teamID), 1) = "M" Then
                    rankNotification = rankNotification & vbCrLf & _
                                       "・次節 " & teamNameOfNextSection(0, 0) & "◯-●" & teamNameOfNextSection(0, 1) & " & " & teamNameOfNextSection(1, 0) & "●-◯" & teamNameOfNextSection(1, 1) & " で、" & vbCrLf & _
                                       "　" & Sheets(seasonName & "_各種記録").Cells(teamID + 2, "BA").Value & "の優勝マジックが消滅…"
                End If
                
            Next teamID
            
            ' 自力V消滅
            For teamID = 0 To 4
                
                If Left(vStatus(1, teamID), 1) <> "自" And Left(vStatus(2, teamID), 1) <> "自" And Left(vStatus(3, teamID), 1) = "自" And Left(vStatus(5, teamID), 1) <> "自" Then
                    rankNotification = rankNotification & vbCrLf & _
                                       "・次節 " & teamNameOfNextSection(0, 0) & "◯-●" & teamNameOfNextSection(0, 1) & " & " & teamNameOfNextSection(1, 0) & "●-◯" & teamNameOfNextSection(1, 1) & " で、" & vbCrLf & _
                                       "　" & Sheets(seasonName & "_各種記録").Cells(teamID + 2, "BA").Value & "の自力優勝が消滅…"
                End If
                
            Next teamID
            
            ' 自力V復活
            For teamID = 0 To 4
                
                If Left(vStatus(1, teamID), 1) = "自" And Left(vStatus(2, teamID), 1) = "自" And Left(vStatus(3, teamID), 1) <> "自" And Left(vStatus(5, teamID), 1) = "自" Then
                    rankNotification = rankNotification & vbCrLf & _
                                       "・次節 " & teamNameOfNextSection(0, 0) & "◯-●" & teamNameOfNextSection(0, 1) & " & " & teamNameOfNextSection(1, 0) & "●-◯" & teamNameOfNextSection(1, 1) & " で、" & vbCrLf & _
                                       "　" & Sheets(seasonName & "_各種記録").Cells(teamID + 2, "BA").Value & "の自力優勝が復活！"
                End If
                
            Next teamID
            
            ' マジック点灯
            For teamID = 0 To 4
                
                If Left(vStatus(1, teamID), 1) <> "M" And Left(vStatus(2, teamID), 1) <> "M" And Left(vStatus(3, teamID), 1) = "M" And Left(vStatus(5, teamID), 1) <> "M" Then
                    rankNotification = rankNotification & vbCrLf & _
                                       "・次節 " & teamNameOfNextSection(0, 0) & "◯-●" & teamNameOfNextSection(0, 1) & " & " & teamNameOfNextSection(1, 0) & "●-◯" & teamNameOfNextSection(1, 1) & " で、" & vbCrLf & _
                                       "　" & Sheets(seasonName & "_各種記録").Cells(teamID + 2, "BA").Value & "の優勝マジックが点灯！"
                End If
                
            Next teamID
            
        End If
        
MPB_NEWS_PATTERN6_END_POINT:

        ' ①●-◯②◯-●
        If vStatus(4, 5) = "true" Then
            ' 優勝
            For teamID = 0 To 4
                
                If Left(vStatus(1, teamID), 1) <> "優" And Left(vStatus(2, teamID), 1) <> "優" And Left(vStatus(4, teamID), 1) = "優" And Left(vStatus(5, teamID), 1) <> "優" Then
                    rankNotification = rankNotification & vbCrLf & _
                                       "・次節 " & teamNameOfNextSection(0, 0) & "●-◯" & teamNameOfNextSection(0, 1) & " & " & teamNameOfNextSection(1, 0) & "◯-●" & teamNameOfNextSection(1, 1) & " で、" & vbCrLf & _
                                       "　" & Sheets(seasonName & "_各種記録").Cells(teamID + 2, "BA").Value & "のリーグ優勝が確定！"
                    GoTo MPB_NEWS_PATTERN7_END_POINT
                End If
                
            Next teamID
            
            ' マジック消滅
            For teamID = 0 To 4
                
                If Left(vStatus(1, teamID), 1) = "M" And Left(vStatus(2, teamID), 1) = "M" And Left(vStatus(4, teamID), 1) <> "M" And Left(vStatus(5, teamID), 1) = "M" Then
                    rankNotification = rankNotification & vbCrLf & _
                                       "・次節 " & teamNameOfNextSection(0, 0) & "●-◯" & teamNameOfNextSection(0, 1) & " & " & teamNameOfNextSection(1, 0) & "◯-●" & teamNameOfNextSection(1, 1) & " で、" & vbCrLf & _
                                       "　" & Sheets(seasonName & "_各種記録").Cells(teamID + 2, "BA").Value & "の優勝マジックが消滅…"
                End If
                
            Next teamID
            
            ' 自力V消滅
            For teamID = 0 To 4
                
                If Left(vStatus(1, teamID), 1) <> "自" And Left(vStatus(2, teamID), 1) <> "自" And Left(vStatus(4, teamID), 1) = "自" And Left(vStatus(5, teamID), 1) <> "自" Then
                    rankNotification = rankNotification & vbCrLf & _
                                       "・次節 " & teamNameOfNextSection(0, 0) & "●-◯" & teamNameOfNextSection(0, 1) & " & " & teamNameOfNextSection(1, 0) & "◯-●" & teamNameOfNextSection(1, 1) & " で、" & vbCrLf & _
                                       "　" & Sheets(seasonName & "_各種記録").Cells(teamID + 2, "BA").Value & "の自力優勝が消滅…"
                End If
                
            Next teamID
            
            ' 自力V復活
            For teamID = 0 To 4
                
                If Left(vStatus(1, teamID), 1) = "自" And Left(vStatus(2, teamID), 1) = "自" And Left(vStatus(4, teamID), 1) <> "自" And Left(vStatus(5, teamID), 1) = "自" Then
                    rankNotification = rankNotification & vbCrLf & _
                                       "・次節 " & teamNameOfNextSection(0, 0) & "●-◯" & teamNameOfNextSection(0, 1) & " & " & teamNameOfNextSection(1, 0) & "◯-●" & teamNameOfNextSection(1, 1) & " で、" & vbCrLf & _
                                       "　" & Sheets(seasonName & "_各種記録").Cells(teamID + 2, "BA").Value & "の自力優勝が復活！"
                End If
                
            Next teamID
            
            ' マジック点灯
            For teamID = 0 To 4
                
                If Left(vStatus(1, teamID), 1) <> "M" And Left(vStatus(2, teamID), 1) <> "M" And Left(vStatus(4, teamID), 1) = "M" And Left(vStatus(5, teamID), 1) <> "M" Then
                    rankNotification = rankNotification & vbCrLf & _
                                       "・次節 " & teamNameOfNextSection(0, 0) & "●-◯" & teamNameOfNextSection(0, 1) & " & " & teamNameOfNextSection(1, 0) & "◯-●" & teamNameOfNextSection(1, 1) & " で、" & vbCrLf & _
                                       "　" & Sheets(seasonName & "_各種記録").Cells(teamID + 2, "BA").Value & "の優勝マジックが点灯！"
                End If
                
            Next teamID
            
        End If
        
MPB_NEWS_PATTERN7_END_POINT:

        ' ①●-◯②●-◯
        If vStatus(5, 5) = "true" Then
            ' 優勝
            For teamID = 0 To 4
                
                If Left(vStatus(1, teamID), 1) <> "優" And Left(vStatus(3, teamID), 1) <> "優" And Left(vStatus(4, teamID), 1) <> "優" And Left(vStatus(5, teamID), 1) = "優" Then
                    rankNotification = rankNotification & vbCrLf & _
                                       "・次節 " & teamNameOfNextSection(0, 0) & "●-◯" & teamNameOfNextSection(0, 1) & " & " & teamNameOfNextSection(1, 0) & "●-◯" & teamNameOfNextSection(1, 1) & " で、" & vbCrLf & _
                                       "　" & Sheets(seasonName & "_各種記録").Cells(teamID + 2, "BA").Value & "のリーグ優勝が確定！"
                    GoTo MPB_NEWS_PATTERN8_END_POINT
                End If
                
            Next teamID
            
            ' マジック消滅
            For teamID = 0 To 4
                
                If Left(vStatus(1, teamID), 1) = "M" And Left(vStatus(3, teamID), 1) = "M" And Left(vStatus(4, teamID), 1) = "M" And Left(vStatus(5, teamID), 1) <> "M" Then
                    rankNotification = rankNotification & vbCrLf & _
                                       "・次節 " & teamNameOfNextSection(0, 0) & "●-◯" & teamNameOfNextSection(0, 1) & " & " & teamNameOfNextSection(1, 0) & "●-◯" & teamNameOfNextSection(1, 1) & " で、" & vbCrLf & _
                                       "　" & Sheets(seasonName & "_各種記録").Cells(teamID + 2, "BA").Value & "の優勝マジックが消滅…"
                End If
                
            Next teamID
            
            ' 自力V消滅
            For teamID = 0 To 4
                
                If Left(vStatus(1, teamID), 1) <> "自" And Left(vStatus(3, teamID), 1) <> "自" And Left(vStatus(4, teamID), 1) <> "自" And Left(vStatus(5, teamID), 1) = "自" Then
                    rankNotification = rankNotification & vbCrLf & _
                                       "・次節 " & teamNameOfNextSection(0, 0) & "●-◯" & teamNameOfNextSection(0, 1) & " & " & teamNameOfNextSection(1, 0) & "●-◯" & teamNameOfNextSection(1, 1) & " で、" & vbCrLf & _
                                       "　" & Sheets(seasonName & "_各種記録").Cells(teamID + 2, "BA").Value & "の自力優勝が消滅…"
                End If
                
            Next teamID
            
            ' 自力V復活
            For teamID = 0 To 4
                
                If Left(vStatus(1, teamID), 1) = "自" And Left(vStatus(3, teamID), 1) = "自" And Left(vStatus(4, teamID), 1) = "自" And Left(vStatus(5, teamID), 1) <> "自" Then
                    rankNotification = rankNotification & vbCrLf & _
                                       "・次節 " & teamNameOfNextSection(0, 0) & "●-◯" & teamNameOfNextSection(0, 1) & " & " & teamNameOfNextSection(1, 0) & "●-◯" & teamNameOfNextSection(1, 1) & " で、" & vbCrLf & _
                                       "　" & Sheets(seasonName & "_各種記録").Cells(teamID + 2, "BA").Value & "の自力優勝が復活！"
                End If
                
            Next teamID
            
            ' マジック点灯
            For teamID = 0 To 4
                
                If Left(vStatus(1, teamID), 1) <> "M" And Left(vStatus(3, teamID), 1) <> "M" And Left(vStatus(4, teamID), 1) <> "M" And Left(vStatus(5, teamID), 1) = "M" Then
                    rankNotification = rankNotification & vbCrLf & _
                                       "・次節 " & teamNameOfNextSection(0, 0) & "●-◯" & teamNameOfNextSection(0, 1) & " & " & teamNameOfNextSection(1, 0) & "●-◯" & teamNameOfNextSection(1, 1) & " で、" & vbCrLf & _
                                       "　" & Sheets(seasonName & "_各種記録").Cells(teamID + 2, "BA").Value & "の優勝マジックが点灯！"
                End If
                
            Next teamID
            
        End If
        
MPB_NEWS_PATTERN8_END_POINT:
        
    End If
    
MPB_NEWS_CHECK_END_POINT:
    
    ' TS/OB枠振り直しのお知らせ
    If numberOfSection = 10 Or numberOfSection = 20 Then
        tsobChangeNotification = tsobChangeNotification & vbCrLf & _
                                 "第" & numberOfSection & "節が終了したので、TS/OB枠の振り直しを行います。" & vbCrLf & _
                                 "- - - - - - - - - -" & vbCrLf & _
                                 "1位:" & Left(Sheets(seasonName & "_各種記録").Cells(2, "B").Value, 1) & "　→　3.5" & vbCrLf & _
                                 "2位:" & Left(Sheets(seasonName & "_各種記録").Cells(3, "B").Value, 1) & "　→　4.0" & vbCrLf & _
                                 "3位:" & Left(Sheets(seasonName & "_各種記録").Cells(4, "B").Value, 1) & "　→　4.5" & vbCrLf & _
                                 "4位:" & Left(Sheets(seasonName & "_各種記録").Cells(5, "B").Value, 1) & "　→　5.0" & vbCrLf & _
                                 "5位:" & Left(Sheets(seasonName & "_各種記録").Cells(6, "B").Value, 1) & "　→　5.5" & vbCrLf & _
                                 "※同率発生時はこの値どおりの変更にならない場合があります。正確な情報をお待ちください。" & vbCrLf & _
                                 "- - - - - - - - - -" & vbCrLf & _
                                 "以上"
    End If
    
    Dim BASIC_CLUSTER_RATIO As Single
    BASIC_CLUSTER_RATIO = 0
    Dim SMALL_CLUSTER_RATIO, BIG_CLUSTER_RATIO As Integer
    SMALL_CLUSTER_RATIO = 70
    BIG_CLUSTER_RATIO = 30
    If SMALL_CLUSTER_RATIO + BIG_CLUSTER_RATIO <> 100 Then
        MsgBox "CLUSTER_RATIOの設定が不正です。"
        End
    End If
    Dim PCR_POSITIVE_RATIO_IN_SMALL_CLUSTER, PCR_POSITIVE_RATIO_IN_BIG_CLUSTER As Single
    PCR_POSITIVE_RATIO_IN_SMALL_CLUSTER = 0.1
    PCR_POSITIVE_RATIO_IN_BIG_CLUSTER = 0.3
    
    Dim BASIC_ACCIDENT_RATIO As Single
    BASIC_ACCIDENT_RATIO = 0.007
    Dim TWO_GAMES_ACCIDENT_RATIO, FIVE_GAMES_ACCIDENT_RATIO, EIGHT_GAMES_ACCIDENT_RATIO, ALL_GAMES_ACCIDENT_RATIO As Integer
    TWO_GAMES_ACCIDENT_RATIO = 70
    FIVE_GAMES_ACCIDENT_RATIO = 15
    EIGHT_GAMES_ACCIDENT_RATIO = 10
    ALL_GAMES_ACCIDENT_RATIO = 5
    If TWO_GAMES_ACCIDENT_RATIO + FIVE_GAMES_ACCIDENT_RATIO + EIGHT_GAMES_ACCIDENT_RATIO + ALL_GAMES_ACCIDENT_RATIO <> 100 Then
        MsgBox "ACCIDENT_RATIOの設定が不正です。"
        End
    End If
    Dim ACCIDENT_PERIOD_SHORT_RATIO, ACCIDENT_PERIOD_NORMAL_RATIO, ACCIDENT_PERIOD_LONG_RATIO As Integer
    ACCIDENT_PERIOD_SHORT_RATIO = 30
    ACCIDENT_PERIOD_NORMAL_RATIO = 40
    ACCIDENT_PERIOD_LONG_RATIO = 30
    If ACCIDENT_PERIOD_SHORT_RATIO + ACCIDENT_PERIOD_NORMAL_RATIO + ACCIDENT_PERIOD_LONG_RATIO <> 100 Then
        MsgBox "ACCIDENT_PERIOD_RATIOの設定が不正です。"
        End
    End If
    
    Dim ACCIDENT_COEFFICIENT_S, ACCIDENT_COEFFICIENT_A, ACCIDENT_COEFFICIENT_B, ACCIDENT_COEFFICIENT_C, _
        ACCIDENT_COEFFICIENT_D, ACCIDENT_COEFFICIENT_E, ACCIDENT_COEFFICIENT_F, ACCIDENT_COEFFICIENT_G, ACCIDENT_COEFFICIENT_n As Single
    ACCIDENT_COEFFICIENT_S = 0.01
    ACCIDENT_COEFFICIENT_A = 0.3
    ACCIDENT_COEFFICIENT_B = 0.5
    ACCIDENT_COEFFICIENT_C = 0.8
    ACCIDENT_COEFFICIENT_D = 1#
    ACCIDENT_COEFFICIENT_E = 1.2
    ACCIDENT_COEFFICIENT_F = 2#
    ACCIDENT_COEFFICIENT_G = 4#
    ACCIDENT_COEFFICIENT_n = 0#
    
    Sheets(seasonName & "_投手データ").Unprotect
    Sheets(seasonName & "_野手データ").Unprotect
    
    For teamID = 0 To 4
        ' 変数の初期化
        teamName = ""
        numberOfGamesPlayedBeforeThisSection = 0
        numberOfGamesPlayedAfterThisSection = 0
        pcrPositiveRate = 0#
        accidentBonusValue = 1#
        pcrResultText = ""
        pcrResultMessage = ""
        
        Sheets(seasonName & "_スケジュール").Activate
        
        ' 試合進行状況の確認
        numberOfGamesPlayedAfterThisSection = ActiveSheet.Cells(2 + numberOfSection, 84 + teamID)
        
        If numberOfSection = 0 Then
            numberOfGamesPlayedBeforeThisSection = -100
        Else
            numberOfGamesPlayedBeforeThisSection = ActiveSheet.Cells(1 + numberOfSection, 84 + teamID)
        End If
        
        Select Case ActiveSheet.Cells(1, 54 + teamID)
            Case Is = "G"
                teamName = "ジャイアンツ"
                accidentBonusValue = 1#
            Case Is = "L"
                teamName = "ライオンズ"
                accidentBonusValue = 1#
            Case Is = "E"
                teamName = "イーグルス"
                accidentBonusValue = 1#
            Case Is = "T"
                teamName = "タイガース"
                accidentBonusValue = 1#
            Case Is = "M"
                teamName = "マリーンズ"
                accidentBonusValue = 1#
            Case Else
                MsgBox "例外が発生しました（1001）"
                End
        End Select
        
        ' スペ判定開始
        If debugModeFlg Then
            MsgBox ActiveSheet.Cells(1, 54 + teamID) & "：判定開始"
        End If
        
        ' クラスター判定
        dice = Rnd()
        If dice < BASIC_CLUSTER_RATIO Then
            dice = Rnd() * 100
            Select Case dice
                Case Is < SMALL_CLUSTER_RATIO
                    pcrPositiveRate = PCR_POSITIVE_RATIO_IN_SMALL_CLUSTER
                Case Is < SMALL_CLUSTER_RATIO + BIG_CLUSTER_RATIO
                    pcrPositiveRate = PCR_POSITIVE_RATIO_IN_BIG_CLUSTER
                Case Else
                    MsgBox "例外が発生しました（1002）"
                    End
            End Select
            pcrResultMessage = teamName & "にてクラスターが発生しました。特例による登録抹消選手は次の通りです。："
            pcrResultText = vbCrLf & _
                            "◇" & teamName & "◇定期スクリーニング検査で球団関係者を含むクラスターが発覚。特例により次の選手が登録抹消となりました。"
        Else
            pcrPositiveRate = 0#
        End If
        
        ' 投手離脱判定
        Sheets(seasonName & "_投手データ").Activate
        pcrResultText = pcrResultText & vbCrLf & _
                        "（投手）"
        For playerID = 4 To 50
            ' 変数の初期化
            rowIdx = teamID * 50 + playerID
            playerName = ActiveSheet.Cells(rowIdx, "B").Value
            playerNameRegistered = ActiveSheet.Cells(rowIdx, "D").Value
            accidentRate = BASIC_ACCIDENT_RATIO * accidentBonusValue
            accidentPeriod = 0
            accidentPeriodString = ""
            accidentOverview = ""
            accidentText = ""
            accidentMessage = ""
            
            If playerName = "" Then
                GoTo STATUS_n_POINT_1
            End If
            If ActiveSheet.Cells(rowIdx, 282 + numberOfGamesPlayedAfterThisSection).Value <> "" Or numberOfGamesPlayedAfterThisSection = numberOfGamesPlayedBeforeThisSection Then
                GoTo ACCIDENT_PERIOD_ZERO_POINT_1
            End If
            ' スペ判定
            ' 基礎スペ率
            Select Case ActiveSheet.Cells(rowIdx, "E").Value
                Case Is = "S"
                    accidentRate = accidentRate * ACCIDENT_COEFFICIENT_S
                Case Is = "A"
                    accidentRate = accidentRate * ACCIDENT_COEFFICIENT_A
                Case Is = "B"
                    accidentRate = accidentRate * ACCIDENT_COEFFICIENT_B
                Case Is = "C"
                    accidentRate = accidentRate * ACCIDENT_COEFFICIENT_C
                Case Is = "D"
                    accidentRate = accidentRate * ACCIDENT_COEFFICIENT_D
                Case Is = "E"
                    accidentRate = accidentRate * ACCIDENT_COEFFICIENT_E
                Case Is = "F"
                    accidentRate = accidentRate * ACCIDENT_COEFFICIENT_F
                Case Is = "G"
                    accidentRate = accidentRate * ACCIDENT_COEFFICIENT_G
                Case Is = "n"
                    accidentRate = accidentRate * ACCIDENT_COEFFICIENT_n
                    GoTo STATUS_n_POINT_1
                Case Else
                    MsgBox "例外が発生しました（1101）"
                    End
            End Select
                    
            ' 救援疲労>=120でスペ率を10倍する処理（投手のみ）
            columnIdxOfStamina = 161 + numberOfGamesPlayedAfterThisSection * 5
            If ActiveSheet.Cells(rowIdx, columnIdxOfStamina).Value - ActiveSheet.Cells(rowIdx, columnIdxOfStamina - 2).Value >= 120 Then
                accidentRate = accidentRate * 10
            End If
            
            ' 試合進行に伴ってスペ率が上昇する処理
            accidentRate = accidentRate * (0.885 + (numberOfGamesPlayedAfterThisSection * 0.01))
              
            ' スペ重さの決定・スペ重さのランダム要素
            dice = Rnd()
            If dice < accidentRate Then
                dice = Rnd() * 100
                Select Case dice
                    Case Is < TWO_GAMES_ACCIDENT_RATIO
                        accidentPeriod = 2
                        accidentPeriodString = "2"
                    Case Is < TWO_GAMES_ACCIDENT_RATIO + FIVE_GAMES_ACCIDENT_RATIO
                        accidentPeriod = 5
                        accidentPeriodString = "5"
                    Case Is < TWO_GAMES_ACCIDENT_RATIO + FIVE_GAMES_ACCIDENT_RATIO + EIGHT_GAMES_ACCIDENT_RATIO
                        accidentPeriod = 8
                        accidentPeriodString = "8"
                    Case Is < TWO_GAMES_ACCIDENT_RATIO + FIVE_GAMES_ACCIDENT_RATIO + EIGHT_GAMES_ACCIDENT_RATIO + ALL_GAMES_ACCIDENT_RATIO
                        accidentPeriod = 24
                        accidentPeriodString = "-"
                        GoTo ACCIDENT_PERIOD_ALL_POINT_1
                End Select
                
                dice = Rnd() * 100
                Select Case dice
                    Case Is < ACCIDENT_PERIOD_SHORT_RATIO
                        accidentPeriod = accidentPeriod - 1
                    Case Is < ACCIDENT_PERIOD_SHORT_RATIO + ACCIDENT_PERIOD_NORMAL_RATIO
                        accidentPeriod = accidentPeriod
                    Case Is < ACCIDENT_PERIOD_SHORT_RATIO + ACCIDENT_PERIOD_NORMAL_RATIO + ACCIDENT_PERIOD_LONG_RATIO
                        accidentPeriod = accidentPeriod + 1
                    Case Else
                        MsgBox "例外が発生しました（1102）"
                        End
                End Select
            Else
                GoTo ACCIDENT_PERIOD_ZERO_POINT_1
            End If
            
ACCIDENT_PERIOD_ALL_POINT_1:
    
            ' 具体的なスペの決定
            dice = Rnd()
            accidentOverview = playerNameRegistered & ":" & Sheets("アクシデント").Cells(Int(5 * dice) + 2, "A").Value & "(" & accidentPeriodString & ")"
            Select Case accidentPeriodString
                Case Is = "2"
                    accidentMessage = playerName & " 選手：" & vbCrLf & _
                                      Sheets("アクシデント").Cells(Int(5 * dice) + 2, "B").Value & vbCrLf & _
                                      Sheets("アクシデント").Cells(12, "B").Value
                    allAccidentResultNotification = allAccidentResultNotification & vbCrLf & _
                                                    "◇" & teamName & "◇" & playerNameRegistered & "選手が" & Sheets("アクシデント").Cells(Int(5 * dice) + 2, "B").Value & Sheets("アクシデント").Cells(12, "B").Value
                Case Is = "5"
                    accidentMessage = playerName & " 選手：" & vbCrLf & _
                                      Sheets("アクシデント").Cells(Int(5 * dice) + 2, "C").Value & vbCrLf & _
                                      Sheets("アクシデント").Cells(12, "C").Value
                    allAccidentResultNotification = allAccidentResultNotification & vbCrLf & _
                                                    "◇" & teamName & "◇" & playerNameRegistered & "選手が" & Sheets("アクシデント").Cells(Int(5 * dice) + 2, "C").Value & Sheets("アクシデント").Cells(12, "C").Value
                Case Is = "8"
                    accidentMessage = playerName & " 選手：" & vbCrLf & _
                                      Sheets("アクシデント").Cells(Int(5 * dice) + 2, "D").Value & vbCrLf & _
                                      Sheets("アクシデント").Cells(12, "D").Value
                    allAccidentResultNotification = allAccidentResultNotification & vbCrLf & _
                                                    "◇" & teamName & "◇" & playerNameRegistered & "選手が" & Sheets("アクシデント").Cells(Int(5 * dice) + 2, "D").Value & Sheets("アクシデント").Cells(12, "D").Value
                Case Is = "-"
                    accidentMessage = playerName & " 選手：" & vbCrLf & _
                                      Sheets("アクシデント").Cells(Int(5 * dice) + 2, "E").Value & vbCrLf & _
                                      Sheets("アクシデント").Cells(12, "E").Value
                    allAccidentResultNotification = allAccidentResultNotification & vbCrLf & _
                                                    "◇" & teamName & "◇" & playerNameRegistered & "選手が" & Sheets("アクシデント").Cells(Int(5 * dice) + 2, "E").Value & Sheets("アクシデント").Cells(12, "E").Value
                Case Else
                    MsgBox "例外が発生しました（1103）"
                    End
            End Select
            
            ' 書き込み
            If debugModeFlg Then
                MsgBox accidentMessage
            End If
            
            For columnIdx = 0 To accidentPeriod - 1
                If 282 + numberOfGamesPlayedAfterThisSection + columnIdx <= 305 Then
                    ActiveSheet.Cells(rowIdx, 282 + numberOfGamesPlayedAfterThisSection + columnIdx).Value = accidentOverview
                End If
            Next columnIdx
            
ACCIDENT_PERIOD_ZERO_POINT_1:
    
            ' 特例判定
            dice = Rnd()
            If dice < pcrPositiveRate Then
                pcrResultMessage = pcrResultMessage & vbCrLf & _
                                   "・" & playerName
                pcrResultText = pcrResultText & playerNameRegistered & " "
                If 282 + numberOfGamesPlayedAfterThisSection <= 305 Then
                    ActiveSheet.Cells(rowIdx, 282 + numberOfGamesPlayedAfterThisSection).Value = playerNameRegistered & ":特例"
                End If
            End If
            
            ' 復帰判定
            If numberOfSection > 0 And ActiveSheet.Cells(rowIdx, 282 + numberOfGamesPlayedBeforeThisSection).Value <> "" And ActiveSheet.Cells(rowIdx, 282 + numberOfGamesPlayedAfterThisSection).Value = "" Then
            
                If debugModeFlg Then
                    MsgBox playerName & " 選手：" & vbCrLf & _
                           "次節からの戦列復帰が首脳陣によって名言されました。"
                End If
                
                allAccidentResultNotification = allAccidentResultNotification & vbCrLf & _
                                                "◇" & teamName & "◇離脱中の" & playerNameRegistered & "選手に関し、次節からの戦列復帰が名言されました。"
            End If
            
STATUS_n_POINT_1:
        
        Next playerID
        
        ' 野手離脱判定
        Sheets(seasonName & "_野手データ").Activate
        pcrResultText = pcrResultText & vbCrLf & _
                        "（野手）"
        For playerID = 4 To 50
            ' 変数の初期化
            rowIdx = teamID * 50 + playerID
            playerName = ActiveSheet.Cells(rowIdx, "B").Value
            playerNameRegistered = ActiveSheet.Cells(rowIdx, "D").Value
            accidentRate = BASIC_ACCIDENT_RATIO
            accidentPeriod = 0
            accidentPeriodString = ""
            accidentOverview = ""
            accidentText = ""
            accidentMessage = ""
            
            If playerName = "" Then
                GoTo STATUS_n_POINT_2
            End If
            If ActiveSheet.Cells(rowIdx, 236 + numberOfGamesPlayedAfterThisSection).Value <> "" Or numberOfGamesPlayedAfterThisSection = numberOfGamesPlayedBeforeThisSection Then
                GoTo ACCIDENT_PERIOD_ZERO_POINT_2
            End If
            ' スペ判定
            ' 基礎スペ率
            Select Case ActiveSheet.Cells(rowIdx, "E").Value
                Case Is = "S"
                    accidentRate = accidentRate * ACCIDENT_COEFFICIENT_S
                Case Is = "A"
                    accidentRate = accidentRate * ACCIDENT_COEFFICIENT_A
                Case Is = "B"
                    accidentRate = accidentRate * ACCIDENT_COEFFICIENT_B
                Case Is = "C"
                    accidentRate = accidentRate * ACCIDENT_COEFFICIENT_C
                Case Is = "D"
                    accidentRate = accidentRate * ACCIDENT_COEFFICIENT_D
                Case Is = "E"
                    accidentRate = accidentRate * ACCIDENT_COEFFICIENT_E
                Case Is = "F"
                    accidentRate = accidentRate * ACCIDENT_COEFFICIENT_F
                Case Is = "G"
                    accidentRate = accidentRate * ACCIDENT_COEFFICIENT_G
                Case Is = "n"
                    accidentRate = accidentRate * ACCIDENT_COEFFICIENT_n
                    GoTo STATUS_n_POINT_2
                Case Else
                    MsgBox "例外が発生しました（1201）"
                    End
            End Select
                    
            ' 試合進行に伴ってスペ率が上昇する処理
            accidentRate = accidentRate * (0.885 + (numberOfGamesPlayedAfterThisSection * 0.01))
            
            ' スペ重さの決定・スペ重さのランダム要素
            dice = Rnd()
            If dice < accidentRate Then
                dice = Rnd() * 100
                Select Case dice
                    Case Is < TWO_GAMES_ACCIDENT_RATIO
                        accidentPeriod = 2
                        accidentPeriodString = "2"
                    Case Is < TWO_GAMES_ACCIDENT_RATIO + FIVE_GAMES_ACCIDENT_RATIO
                        accidentPeriod = 5
                        accidentPeriodString = "5"
                    Case Is < TWO_GAMES_ACCIDENT_RATIO + FIVE_GAMES_ACCIDENT_RATIO + EIGHT_GAMES_ACCIDENT_RATIO
                        accidentPeriod = 8
                        accidentPeriodString = "8"
                    Case Is < TWO_GAMES_ACCIDENT_RATIO + FIVE_GAMES_ACCIDENT_RATIO + EIGHT_GAMES_ACCIDENT_RATIO + ALL_GAMES_ACCIDENT_RATIO
                        accidentPeriod = 24
                        accidentPeriodString = "-"
                        GoTo ACCIDENT_PERIOD_ALL_POINT_2
                End Select
                
                dice = Rnd() * 100
                Select Case dice
                    Case Is < ACCIDENT_PERIOD_SHORT_RATIO
                        accidentPeriod = accidentPeriod - 1
                    Case Is < ACCIDENT_PERIOD_SHORT_RATIO + ACCIDENT_PERIOD_NORMAL_RATIO
                        accidentPeriod = accidentPeriod
                    Case Is < ACCIDENT_PERIOD_SHORT_RATIO + ACCIDENT_PERIOD_NORMAL_RATIO + ACCIDENT_PERIOD_LONG_RATIO
                        accidentPeriod = accidentPeriod + 1
                    Case Else
                        MsgBox "例外が発生しました（1202）"
                        End
                End Select
            Else
                GoTo ACCIDENT_PERIOD_ZERO_POINT_2
            End If
            
ACCIDENT_PERIOD_ALL_POINT_2:
    
            ' 具体的なスペの決定
            dice = Rnd()
            accidentOverview = playerNameRegistered & ":" & Sheets("アクシデント").Cells(Int(5 * dice) + 2, "A").Value & "(" & accidentPeriodString & ")"
            Select Case accidentPeriodString
                Case Is = "2"
                    accidentMessage = playerName & " 選手：" & vbCrLf & _
                                      Sheets("アクシデント").Cells(Int(4 * dice) + 8, "B").Value & vbCrLf & _
                                      Sheets("アクシデント").Cells(12, "B").Value
                    allAccidentResultNotification = allAccidentResultNotification & vbCrLf & _
                                                    "◇" & teamName & "◇" & playerNameRegistered & "選手が" & Sheets("アクシデント").Cells(Int(5 * dice) + 2, "B").Value & Sheets("アクシデント").Cells(12, "B").Value
                Case Is = "5"
                    accidentMessage = playerName & " 選手：" & vbCrLf & _
                                      Sheets("アクシデント").Cells(Int(4 * dice) + 8, "C").Value & vbCrLf & _
                                      Sheets("アクシデント").Cells(12, "C").Value
                    allAccidentResultNotification = allAccidentResultNotification & vbCrLf & _
                                                    "◇" & teamName & "◇" & playerNameRegistered & "選手が" & Sheets("アクシデント").Cells(Int(5 * dice) + 2, "C").Value & Sheets("アクシデント").Cells(12, "C").Value
                Case Is = "8"
                    accidentMessage = playerName & " 選手：" & vbCrLf & _
                                      Sheets("アクシデント").Cells(Int(4 * dice) + 8, "D").Value & vbCrLf & _
                                      Sheets("アクシデント").Cells(12, "D").Value
                    allAccidentResultNotification = allAccidentResultNotification & vbCrLf & _
                                                    "◇" & teamName & "◇" & playerNameRegistered & "選手が" & Sheets("アクシデント").Cells(Int(5 * dice) + 2, "D").Value & Sheets("アクシデント").Cells(12, "D").Value
                Case Is = "-"
                    accidentMessage = playerName & " 選手：" & vbCrLf & _
                                      Sheets("アクシデント").Cells(Int(4 * dice) + 8, "E").Value & vbCrLf & _
                                      Sheets("アクシデント").Cells(12, "E").Value
                    allAccidentResultNotification = allAccidentResultNotification & vbCrLf & _
                                                    "◇" & teamName & "◇" & playerNameRegistered & "選手が" & Sheets("アクシデント").Cells(Int(5 * dice) + 2, "E").Value & Sheets("アクシデント").Cells(12, "E").Value
                Case Else
                    MsgBox "例外が発生しました（1203）"
                    End
            End Select
            
            ' 書き込み
            If debugModeFlg Then
                MsgBox accidentMessage
            End If
            
            For columnIdx = 0 To accidentPeriod - 1
                If 236 + numberOfGamesPlayedAfterThisSection + columnIdx <= 259 Then
                    ActiveSheet.Cells(rowIdx, 236 + numberOfGamesPlayedAfterThisSection + columnIdx).Value = accidentOverview
                End If
            Next columnIdx
            
ACCIDENT_PERIOD_ZERO_POINT_2:
    
            ' 特例判定
            dice = Rnd()
            If dice < pcrPositiveRate Then
                pcrResultMessage = pcrResultMessage & vbCrLf & _
                                   "・" & playerName
                pcrResultText = pcrResultText & playerNameRegistered & " "
                If 236 + numberOfGamesPlayedAfterThisSection <= 259 Then
                    ActiveSheet.Cells(rowIdx, 236 + numberOfGamesPlayedAfterThisSection).Value = playerNameRegistered & ":特例"
                End If
            End If
            
            ' 復帰判定
            If numberOfSection > 0 And ActiveSheet.Cells(rowIdx, 236 + numberOfGamesPlayedBeforeThisSection).Value <> "" And ActiveSheet.Cells(rowIdx, 236 + numberOfGamesPlayedAfterThisSection).Value = "" Then
                
                If debugModeFlg Then
                    MsgBox playerName & " 選手：" & vbCrLf & _
                           "次節からの戦列復帰が首脳陣によって名言されました。"
                End If
                
                allAccidentResultNotification = allAccidentResultNotification & vbCrLf & _
                                                "◇" & teamName & "◇離脱中の" & playerNameRegistered & "選手に関し、次節からの戦列復帰が名言されました。"
            End If
            
STATUS_n_POINT_2:
    
        Next playerID
            
        ' クラスター判定の結果を出力
        If pcrPositiveRate > 0 Then
            ' MsgBox pcrResultMessage
            allAccidentResultNotification = allAccidentResultNotification & pcrResultText
        End If
        
        ' 終了の出力
        Sheets(seasonName & "_スケジュール").Activate
        If debugModeFlg Then
            If numberOfGamesPlayedAfterThisSection > numberOfGamesPlayedBeforeThisSection Then
                MsgBox ActiveSheet.Cells(1, 54 + teamID) & "：判定終了"
            Else
                MsgBox ActiveSheet.Cells(1, 54 + teamID) & "：判定終了(特例のみ)"
            End If
        End If
    Next teamID
    
    Sheets(seasonName & "_投手データ").Protect AllowFormattingColumns:=True, AllowFormattingRows:=True
    Sheets(seasonName & "_野手データ").Protect AllowFormattingColumns:=True, AllowFormattingRows:=True
    Sheets(seasonName & "_スケジュール").Activate
    
    If Not debugModeFlg Then
        Call バックアップ
        
        Open "C:\Users\TaiNo\マイドライブ\純正リアタイ部_送信待機\accident.txt" For Output As #1
            Print #1, allAccidentResultNotification;
        Close #1
        
        If rankNotification <> "【MPBニュース】" Then
            Open "C:\Users\TaiNo\マイドライブ\純正リアタイ部_送信待機\news.txt" For Output As #3
                Print #3, rankNotification;
            Close #3
        End If
        
        If tsobChangeNotification <> "【TS/OB枠振り直しのお知らせ】" Then
            Open "C:\Users\TaiNo\マイドライブ\純正リアタイ部_送信待機\tsob.txt" For Output As #4
                Print #4, tsobChangeNotification;
            Close #4
        End If
        
        Call 画像保存
        
        Application.ScreenUpdating = True
        ActiveWorkbook.Close SaveChanges:=True
    Else
        MsgBox allAccidentResultNotification
        MsgBox rankNotification
        MsgBox tsobChangeNotification

        MsgBox "処理を終了します"
    End If
    
End Sub


