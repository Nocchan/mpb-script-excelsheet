Attribute VB_Name = "utils"
Option Explicit

Public isDebugMode As Boolean

' Debugレベルのメッセージ
Public Function showMessageDebug(message As String, Optional title As String = "")
    
    If Not isDebugMode Then
        Exit Function
    End If
    
    MsgBox message, _
           vbInformation, _
           "[DEBUG] " & title

End Function

' Infoレベルのメッセージ
Public Function showMessageInfo(message As String, Optional title As String = "")

    MsgBox message, _
           vbExclamation, _
           "[INFO] " & title

End Function

' Errorレベルのメッセージ
Public Function showMessageError(message As String, Optional title As String = "")

    MsgBox message, _
           vbCritical, _
           "[ERROR] " & title

End Function

' バックアップを作成
Public Function makeBackupFile()
    
    If isDebugMode Then
        ActiveWorkbook.SaveCopyAs Filename:="C:\Users\TaiNo\Desktop\ペナントバックアップ\" & Format(Now, "yyyymmddhhnnss") & "-Debug.xlsm"
    Else
        ActiveWorkbook.SaveCopyAs Filename:="C:\Users\TaiNo\Desktop\ペナントバックアップ\" & Format(Now, "yyyymmddhhnnss") & ".xlsm"
    End If
    
End Function

' デバッグモードでの起動に切り替え
Public Function enableDebugMode()

    isDebugMode = True
    Call showMessageInfo("デバッグモードで起動", "enableDebugMode")

End Function

' スケジュールシートからの呼び出しかを判定
Public Function isScheduleSheet() As Boolean

    isScheduleSheet = (ActiveSheet.Name = ActiveSheet.Cells(1, "A").Value & "_スケジュール")

End Function

' 投手/野手データシートからの呼び出しかを判定
Public Function isSeasonDataSheet() As Boolean
    
    isSeasonDataSheet = (ActiveSheet.Name = ActiveSheet.Cells(1, "H").Value & "_投手データ") Or (ActiveSheet.Name = ActiveSheet.Cells(1, "H").Value & "_野手データ")
    
End Function

' 記録室シートからの呼び出しかを判定
Public Function isCareerDataSheet() As Boolean
    
    isCareerDataSheet = (ActiveSheet.Name = "記録室_" & ActiveSheet.Cells(2, "A").Value)
    
End Function

' 改行して追記
Public Function addLineToText(baseText As String, addText As String)

    addLineToText = baseText & vbCrLf & _
                 addText

End Function

' 抽選
Public Function drawFromDictionary(dict As Dictionary)
    
    Dim dKey
    Dim dVal As Single
    Dim dice As Single
    
    dVal = 0
    Randomize
    dice = Rnd() * WorksheetFunction.Sum(dict.Items)
    
    For Each dKey In dict.Keys
        dVal = dVal + dict.Item(dKey)
        
        If dice < dVal Then
            drawFromDictionary = dKey
            Exit Function
        End If
    Next
    
End Function

' テキストファイルの出力
Public Function saveTxtFile(text As String, path As String)

    Dim fileNumber As Integer
    fileNumber = FreeFile
    
    Open path For Output As fileNumber
        Print #fileNumber, text;
    Close fileNumber

End Function

' 画像ファイルの出力
Public Function savePngFile(pictureRange As Range, path As String)
    
    Dim pictureRangeTmp As ChartObject
    Dim minFileSize As Long
    
    pictureRange.CopyPicture
    Set pictureRangeTmp = Sheets("(tmp)").ChartObjects.Add(0, 0, pictureRange.Width, pictureRange.Height)
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
