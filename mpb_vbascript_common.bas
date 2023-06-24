Attribute VB_Name = "mpb_vbascript_common"
Option Explicit

Public debugModeFlg As Boolean

' Debugレベルのメッセージ
Public Function MessageDebug(message As String, Optional title As String = "")
    
    If Not debugModeFlg Then
        Exit Function
    End If
    
    MsgBox message, _
           vbInformation, _
           "[DEBUG] " & title

End Function

' Infoレベルのメッセージ
Public Function MessageInfo(message As String, Optional title As String = "")

    MsgBox message, _
           vbExclamation, _
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

' 改行して追記
Public Function AddRowText(baseText As String, addText As String)

    AddRowText = baseText & vbCrLf & _
                 addText

End Function

' 抽選
Public Function DrawFromDict(dict As Dictionary)
    
    Dim dKey
    Dim dVal As Single
    Dim dice As Single
    
    dVal = 0
    Randomize
    dice = Rnd() * WorksheetFunction.Sum(dict.Items)
    
    For Each dKey In dict.Keys
        dVal = dVal + dict.Item(dKey)
        
        If dice < dVal Then
            DrawFromDict = dKey
            Exit Function
        End If
    Next
    
End Function

' テキストファイルの出力
Public Function OutputText(text As String, path As String)

    Dim fileNumber As Integer
    fileNumber = FreeFile
    
    Open path For Output As fileNumber
        Print #fileNumber, text;
    Close fileNumber

End Function

' 画像ファイルの出力
Public Function OutputPicture(pictureRange As Range, path As String)
    
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
