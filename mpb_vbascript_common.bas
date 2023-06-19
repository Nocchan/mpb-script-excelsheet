Attribute VB_Name = "mpb_vbascript_common"
Public debugModeFlg As Boolean

' バックアップファイルを作成
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
    MsgBox "デバッグモードで起動", vbInformation

End Function

' 記録室シートからの呼び出しかを判定
Public Function IsCareerDataSheet() As Boolean
    
    IsCareerDataSheet = (ActiveSheet.Name = "記録室_" & ActiveSheet.Cells(2, "A").Value)
    
End Function

' スケジュールシートからの呼び出しかを判定
Public Function IsScheduleSheet() As Boolean

    IsScheduleSheet = (ActiveSheet.Name = ActiveSheet.Cells(1, "A").Value & "_スケジュール")

End Function

' 投手/野手データシートからの呼び出しかを判定
Public Function IsSeasonDataSheet() As Boolean
    
    IsSeasonDataSheet = (ActiveSheet.Name = ActiveSheet.Cells(1, "H").Value & "_投手データ") Or (ActiveSheet.Name = ActiveSheet.Cells(1, "H").Value & "_野手データ")
    
End Function
