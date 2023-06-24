Attribute VB_Name = "mpb_vbascript_common"
Option Explicit

Public debugModeFlg As Boolean

' Debug���x���̃��b�Z�[�W
Public Function MessageDebug(message As String, Optional title As String = "")
    
    If Not debugModeFlg Then
        Exit Function
    End If
    
    MsgBox message, _
           vbInformation, _
           "[DEBUG] " & title

End Function

' Info���x���̃��b�Z�[�W
Public Function MessageInfo(message As String, Optional title As String = "")

    MsgBox message, _
           vbExclamation, _
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

' ���s���ĒǋL
Public Function AddRowText(baseText As String, addText As String)

    AddRowText = baseText & vbCrLf & _
                 addText

End Function

' ���I
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

' �e�L�X�g�t�@�C���̏o��
Public Function OutputText(text As String, path As String)

    Dim fileNumber As Integer
    fileNumber = FreeFile
    
    Open path For Output As fileNumber
        Print #fileNumber, text;
    Close fileNumber

End Function

' �摜�t�@�C���̏o��
Public Function OutputPicture(pictureRange As Range, path As String)
    
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
