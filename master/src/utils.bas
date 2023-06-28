Attribute VB_Name = "utils"
Option Explicit

Public isDebugMode As Boolean

' Debug���x���̃��b�Z�[�W
Public Function showMessageDebug(message As String, Optional title As String = "")
    
    If Not isDebugMode Then
        Exit Function
    End If
    
    MsgBox message, _
           vbInformation, _
           "[DEBUG] " & title

End Function

' Info���x���̃��b�Z�[�W
Public Function showMessageInfo(message As String, Optional title As String = "")

    MsgBox message, _
           vbExclamation, _
           "[INFO] " & title

End Function

' Error���x���̃��b�Z�[�W
Public Function showMessageError(message As String, Optional title As String = "")

    MsgBox message, _
           vbCritical, _
           "[ERROR] " & title

End Function

' �o�b�N�A�b�v���쐬
Public Function makeBackupFile()
    
    If isDebugMode Then
        ActiveWorkbook.SaveCopyAs Filename:="C:\Users\TaiNo\Desktop\�y�i���g�o�b�N�A�b�v\" & Format(Now, "yyyymmddhhnnss") & "-Debug.xlsm"
    Else
        ActiveWorkbook.SaveCopyAs Filename:="C:\Users\TaiNo\Desktop\�y�i���g�o�b�N�A�b�v\" & Format(Now, "yyyymmddhhnnss") & ".xlsm"
    End If
    
End Function

' �f�o�b�O���[�h�ł̋N���ɐ؂�ւ�
Public Function enableDebugMode()

    isDebugMode = True
    Call showMessageInfo("�f�o�b�O���[�h�ŋN��", "enableDebugMode")

End Function

' �X�P�W���[���V�[�g����̌Ăяo�����𔻒�
Public Function isScheduleSheet() As Boolean

    isScheduleSheet = (ActiveSheet.Name = ActiveSheet.Cells(1, "A").Value & "_�X�P�W���[��")

End Function

' ����/���f�[�^�V�[�g����̌Ăяo�����𔻒�
Public Function isSeasonDataSheet() As Boolean
    
    isSeasonDataSheet = (ActiveSheet.Name = ActiveSheet.Cells(1, "H").Value & "_����f�[�^") Or (ActiveSheet.Name = ActiveSheet.Cells(1, "H").Value & "_���f�[�^")
    
End Function

' �L�^���V�[�g����̌Ăяo�����𔻒�
Public Function isCareerDataSheet() As Boolean
    
    isCareerDataSheet = (ActiveSheet.Name = "�L�^��_" & ActiveSheet.Cells(2, "A").Value)
    
End Function

' ���s���ĒǋL
Public Function addLineToText(baseText As String, addText As String)

    addLineToText = baseText & vbCrLf & _
                 addText

End Function

' ���I
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

' �e�L�X�g�t�@�C���̏o��
Public Function saveTxtFile(text As String, path As String)

    Dim fileNumber As Integer
    fileNumber = FreeFile
    
    Open path For Output As fileNumber
        Print #fileNumber, text;
    Close fileNumber

End Function

' �摜�t�@�C���̏o��
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
