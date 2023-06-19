Attribute VB_Name = "mpb_vbascript_common"
Public debugModeFlg As Boolean

' �o�b�N�A�b�v�t�@�C�����쐬
Public Function Backup()
    
    If debugModeFlg Then
        ActiveWorkbook.SaveCopyAs Filename:="C:\Users\TaiNo\Desktop\�y�i���g�o�b�N�A�b�v\" & Format(Now, "yyyymmddhhnnss") & "-Debug.xlsm"
    Else
        ActiveWorkbook.SaveCopyAs Filename:="C:\Users\TaiNo\Desktop\�y�i���g�o�b�N�A�b�v\" & Format(Now, "yyyymmddhhnnss") & ".xlsm"
    End If
    
End Function

' �f�o�b�O���[�h�ł̋N���ɐ؂�ւ�
Public Function DebugMode()

    debugModeFlg = True
    MsgBox "�f�o�b�O���[�h�ŋN��", vbInformation

End Function

' �L�^���V�[�g����̌Ăяo�����𔻒�
Public Function IsCareerDataSheet() As Boolean
    
    IsCareerDataSheet = (ActiveSheet.Name = "�L�^��_" & ActiveSheet.Cells(2, "A").Value)
    
End Function

' �X�P�W���[���V�[�g����̌Ăяo�����𔻒�
Public Function IsScheduleSheet() As Boolean

    IsScheduleSheet = (ActiveSheet.Name = ActiveSheet.Cells(1, "A").Value & "_�X�P�W���[��")

End Function

' ����/���f�[�^�V�[�g����̌Ăяo�����𔻒�
Public Function IsSeasonDataSheet() As Boolean
    
    IsSeasonDataSheet = (ActiveSheet.Name = ActiveSheet.Cells(1, "H").Value & "_����f�[�^") Or (ActiveSheet.Name = ActiveSheet.Cells(1, "H").Value & "_���f�[�^")
    
End Function
