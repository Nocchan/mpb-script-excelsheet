Attribute VB_Name = "mpb_vbascript_common"
Public debugModeFlg As Boolean

Public Function Backup()
    
    If debugModeFlg Then
        ActiveWorkbook.SaveCopyAs Filename:="C:\Users\TaiNo\Desktop\ペナントバックアップ\" & Format(Now, "yyyymmddhhnnss") & "-Debug.xlsm"
    Else
        ActiveWorkbook.SaveCopyAs Filename:="C:\Users\TaiNo\Desktop\ペナントバックアップ\" & Format(Now, "yyyymmddhhnnss") & ".xlsm"
    End If
    
End Function

Public Function DebugMode()

    debugModeFlg = True
    MsgBox "デバッグモードで起動"

End Function

Public Function IsCareerDataSheet() As Boolean
    
    IsCareerDataSheet = (ActiveSheet.Name = "記録室_" & ActiveSheet.Cells(2, "A").Value)
    
End Function

Public Function IsScheduleSheet() As Boolean

    IsScheduleSheet = (ActiveSheet.Name = ActiveSheet.Cells(1, "A").Value & "_スケジュール")

End Function

Public Function IsSeasonDataSheet() As Boolean
    
    IsSeasonDataSheet = (ActiveSheet.Name = ActiveSheet.Cells(1, "H").Value & "_投手データ") Or (ActiveSheet.Name = ActiveSheet.Cells(1, "H").Value & "_野手データ")
    
End Function
