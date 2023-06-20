Attribute VB_Name = "mpb_vbascript_common"
Public debugModeFlg As Boolean

' Infoレベルのメッセージ
Public Function MessageInfo(message As String, Optional title As String = "")

    MsgBox message, _
           vbInformation, _
           "[INFO] " & title

End Function

' Errorレベルのメッセージ
Public Function MessageError(message As String, Optional title As String = "")

    MsgBox message, _
           vbCritical, _
           "[ERROR] " & title

End Function

' バックアップを作成
Public Function Backup()
    
    If debugModeFlg Then
        ActiveWorkbook.SaveCopyAs Filename:="C:\Users\TaiNo\Desktop\ペナントバックアップ\" & Format(Now, "yyyymmddhhnnss") & "-Debug.xlsm"
    Else
        ActiveWorkbook.SaveCopyAs Filename:="C:\Users\TaiNo\Desktop\ペナントバックアップ\" & Format(Now, "yyyymmddhhnnss") & ".xlsm"
    End If
    
End Function

' デバッグモードでの起動に切り替え
Public Function DebugMode()

    debugModeFlg = True
    Call MessageInfo("デバッグモードで起動", "DebugMode")

End Function

' スケジュールシートからの呼び出しかを判定
Public Function IsScheduleSheet() As Boolean

    IsScheduleSheet = (ActiveSheet.Name = ActiveSheet.Cells(1, "A").Value & "_スケジュール")

End Function

' 投手/野手データシートからの呼び出しかを判定
Public Function IsSeasonDataSheet() As Boolean
    
    IsSeasonDataSheet = (ActiveSheet.Name = ActiveSheet.Cells(1, "H").Value & "_投手データ") Or (ActiveSheet.Name = ActiveSheet.Cells(1, "H").Value & "_野手データ")
    
End Function

' 記録室シートからの呼び出しかを判定
Public Function IsCareerDataSheet() As Boolean
    
    IsCareerDataSheet = (ActiveSheet.Name = "記録室_" & ActiveSheet.Cells(2, "A").Value)
    
End Function

' テキストファイルの出力
Public Function outputText(text As String, path As String)

    Dim fileNumber As Integer
    fileNumber = FreeFile
    
    Open path For Output As fileNumber
        Print fileNumber, text;
    Close fileNumber

End Function

' 画像ファイルの出力
Public Function outputPicture(pictureRange As range, path As String)
    
    If Dir(path) <> "" Then
        Call MessageError("画像配置不可エラー", "outputPicture")
        End
    End If
    
    Dim pictureRangeTmp As ChartObject
    Dim minFileSize As Long
    
    pictureRange.CopyPicture
    Set pictureRangeTmp = Sheets("アクシデント").ChartObjects.Add(0, 0, pictureRange.Width, pictureRange.Height)
    pictureRangeTmp.Chart.Export path
    minFileSize = FileLen(path)
    
    Do Until FileLen(path) > minFileSize
        pictureRangeTmp.Chart.Paste
        pictureRangeTmp.Chart.Export path
        DoEvents
    Loop
    
    pictureRangeTmp.Delete
    Set pictureRangeTmp = Nothing

End Function
