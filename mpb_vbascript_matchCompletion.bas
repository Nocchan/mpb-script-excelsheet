Attribute VB_Name = "mpb_vbascript_matchCompletion"
Dim strSeason As String
Dim numSection As Integer

Dim pictureRangeSchedule, pictureRangeRanking As ChartObject
Dim pictureName As String
Dim minFileSize As Long

Sub matchCompletion()
    
    ' デバッグモード確認
    ' Call DebugMode

    ' エラーチェック
    If Not IsScheduleSheet() Then
        MsgBox "matchCompletion.Error : 0000"
        End
    End If
    
    ' 初期化
    Call Initialize
    
    ' スケジュールのステータスチェック
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



End Function

' 終了時処理
Function ExitProcess()



End Function


' 節が完了してスペ判定を行える状態かを判定
Function IsSectionCompleted() As Boolean

    

End Function

' 節の進行により発生する、あらかじめ予定されているイベントを出力
Function MakeMPBNewsSeasonEvent()



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

