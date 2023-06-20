Attribute VB_Name = "mpb_vbascript_common"
Public debugModeFlg As Boolean

' Info���x���̃��b�Z�[�W
Public Function MessageInfo(message As String, Optional title As String = "")

    MsgBox message, _
           vbInformation, _
           "[INFO] " & title

End Function

' Error���x���̃��b�Z�[�W
Public Function MessageError(message As String, Optional title As String = "")

    MsgBox message, _
           vbCritical, _
           "[ERROR] " & title

End Function

' �o�b�N�A�b�v���쐬
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
    Call MessageInfo("�f�o�b�O���[�h�ŋN��", "DebugMode")

End Function

' �X�P�W���[���V�[�g����̌Ăяo�����𔻒�
Public Function IsScheduleSheet() As Boolean

    IsScheduleSheet = (ActiveSheet.Name = ActiveSheet.Cells(1, "A").Value & "_�X�P�W���[��")

End Function

' ����/���f�[�^�V�[�g����̌Ăяo�����𔻒�
Public Function IsSeasonDataSheet() As Boolean
    
    IsSeasonDataSheet = (ActiveSheet.Name = ActiveSheet.Cells(1, "H").Value & "_����f�[�^") Or (ActiveSheet.Name = ActiveSheet.Cells(1, "H").Value & "_���f�[�^")
    
End Function

' �L�^���V�[�g����̌Ăяo�����𔻒�
Public Function IsCareerDataSheet() As Boolean
    
    IsCareerDataSheet = (ActiveSheet.Name = "�L�^��_" & ActiveSheet.Cells(2, "A").Value)
    
End Function

' �e�L�X�g�t�@�C���̏o��
Public Function outputText(text As String, path As String)

    Dim fileNumber As Integer
    fileNumber = FreeFile
    
    Open path For Output As fileNumber
        Print fileNumber, text;
    Close fileNumber

End Function

' �摜�t�@�C���̏o��
Public Function outputPicture(pictureRange As range, path As String)
    
    If Dir(path) <> "" Then
        Call MessageError("�摜�z�u�s�G���[", "outputPicture")
        End
    End If
    
    Dim pictureRangeTmp As ChartObject
    Dim minFileSize As Long
    
    pictureRange.CopyPicture
    Set pictureRangeTmp = Sheets("�A�N�V�f���g").ChartObjects.Add(0, 0, pictureRange.Width, pictureRange.Height)
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
