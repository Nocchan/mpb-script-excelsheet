Attribute VB_Name = "mpb_vbascript_common"
Public debugModeFlg As Boolean

Public Function Backup()
    
    If debugModeFlg Then
        ActiveWorkbook.SaveCopyAs Filename:="C:\Users\TaiNo\Desktop\�y�i���g�o�b�N�A�b�v\" & Format(Now, "yyyymmddhhnnss") & "-Debug.xlsm"
    Else
        ActiveWorkbook.SaveCopyAs Filename:="C:\Users\TaiNo\Desktop\�y�i���g�o�b�N�A�b�v\" & Format(Now, "yyyymmddhhnnss") & ".xlsm"
    End If
    
End Function

Public Function DebugMode()

    debugModeFlg = True
    MsgBox "�f�o�b�O���[�h�ŋN��"

End Function

Public Function IsCareerDataSheet() As Boolean
    
    IsCareerDataSheet = (ActiveSheet.Name = "�L�^��_" & ActiveSheet.Cells(2, "A").Value)
    
End Function

Public Function IsScheduleSheet() As Boolean

    IsScheduleSheet = (ActiveSheet.Name = ActiveSheet.Cells(1, "A").Value & "_�X�P�W���[��")

End Function

Public Function IsSeasonDataSheet() As Boolean
    
    IsSeasonDataSheet = (ActiveSheet.Name = ActiveSheet.Cells(1, "H").Value & "_����f�[�^") Or (ActiveSheet.Name = ActiveSheet.Cells(1, "H").Value & "_���f�[�^")
    
End Function
