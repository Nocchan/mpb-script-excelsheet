Attribute VB_Name = "mpb_vbascript_matchCompletion"
Dim season As String
Dim game As Integer
Dim section As Integer

Dim DICT_TEAMNAME As Object
Dim DICT_ACCIDENT_HDCP As Object

Sub matchCompletion()
    
    ' �f�o�b�O���[�h
    ' Call DebugMode

    ' �ďo���m�F
    If Not IsScheduleSheet() Then
        MsgBox "�ďo���m�F�G���[", _
               vbCritical, _
               "[ERROR] matchCompletion"
        End
    End If
    
    Call Initialize
    
    If IsSectionCompleted() Then
        Call MakeMPBNewsSeasonEvent
        Call MakeMPBNewsOfThisSection
        Call MakeMPBNewsOfAccident
        Call MakeMPBNewsOfNextGame
    End If
    
    Call SavePictureOfSchedule
    Call SavePictureOfRanking
    
    Call ExitProcess

End Sub

' �萔�E�V�[�g��Ԃ̏�����
Function Initialize()

    If Not debugModeFlg Then
        Application.ScreenUpdating = False
    End If
    
    Application.Calculate
    
    season = ActiveSheet.Cells(1, "A").Value
    game = WorksheetFunction.CountIf(ActiveSheet.Range("BA2:BA241"), 0) / 4
    section = WorksheetFunction.CountIf(ActiveSheet.Range("BA2:BA241"), 0) / 8
    
    Sheets(season & "_����f�[�^").Unprotect
    Sheets(season & "_���f�[�^").Unprotect
    
    Set DICT_TEAMNAME = CreateObject("Scripting.Dictionary")
    With DICT_TEAMNAME
        .Add "G", "�W���C�A���c"
        .Add "M", "�}���[���Y"
        .Add "T", "�^�C�K�[�X"
        .Add "L", "���C�I���Y"
        .Add "E", "�C�[�O���X"
    End With
    
    Set DICT_ACCIDENT_HDCP = CreateObject("Scripting.Dictionary")
    With DICT_ACCIDENT_HDCP
        .Add "G", 1#
        .Add "M", 1#
        .Add "T", 1#
        .Add "L", 1#
        .Add "E", 1#
    End With
    
End Function

' �I��������
Function ExitProcess()

    Sheets(season & "_����f�[�^").Protect AllowFormattingColumns:=True, AllowFormattingRows:=True
    Sheets(season & "_���f�[�^").Protect AllowFormattingColumns:=True, AllowFormattingRows:=True

    If Not debugModeFlg Then
        Application.ScreenUpdating = True
    End If
    
    End
    
End Function

' �߂��������ăX�y������s�����Ԃ��𔻒�
Function IsSectionCompleted() As Boolean

    ' ��������������߂��������Ă��Ȃ����Ƃ��킩��p�^�[��
    If game <> section * 2 Then
        IsSectionCompleted = False
        Exit Function
    End If
    
    ' �߂͊������Ă��邪�s�����͂�����p�^�[��
    If ActiveSheet.Cells(section * 8 + 3, "D").Value <> "" Or ActiveSheet.Cells(section * 8 + 7, "D").Value <> "" Or _
       ActiveSheet.Cells(section * 8 + 3, "F").Value <> "" Or ActiveSheet.Cells(section * 8 + 7, "F").Value <> "" Or _
       ActiveSheet.Cells(section * 8 + 3, "H").Value <> "" Or ActiveSheet.Cells(section * 8 + 7, "H").Value <> "" Then
        MsgBox "�s�����̓G���[", _
               vbCritical, _
               "[ERROR] IsSectionCompleted"
        Call ExitProcess
    End If
    
    ' �\���攭���o�����Ă��Ȃ��p�^�[��
    If section > 0 Then
        IsSectionCompleted = True
        Exit Function
    End If
    
    If ActiveSheet.Cells(section * 8 + 2, "D").Value = "" Or ActiveSheet.Cells(section * 8 + 6, "D").Value = "" Or _
       ActiveSheet.Cells(section * 8 + 2, "H").Value = "" Or ActiveSheet.Cells(section * 8 + 6, "H").Value = "" Then
        MsgBox "�\���攭�������G���[", _
               vbCritical, _
               "[ERROR] IsSectionCompleted"
        Call ExitProcess
    End If
    
    IsSectionCompleted = True
    
End Function

' �߂̐i�s�ɂ�蔭������A���炩���ߗ\�肳��Ă���C�x���g���o��
Function MakeMPBNewsSeasonEvent()



End Function

' �߂̐i�s�ɂ�蔭������A�D���}�W�b�N�⎩�͗D���Ɋւ���C�x���g���o��
Function MakeMPBNewsOfThisSection()



End Function

' �X�y����E���ʂ��o��
Function MakeMPBNewsOfAccident()



End Function

' ���ߓ��������̈˗����o��
Function MakeMPBNewsOfNextGame()



End Function

' �X�P�W���[���摜���o��
Function SavePictureOfSchedule()



End Function

' ���щ摜���o��
Function SavePictureOfRanking()



End Function


Sub �摜�ۑ�()

    ' �G���[�`�F�b�N
    If ActiveSheet.Name <> ActiveSheet.Cells(1, "A").Value & "_�X�P�W���[��" Then
        MsgBox "�V�[�g���܂���A1�Z���̃V�[�Y���w�肪�s���ł��B"
        End
    End If

    Application.ScreenUpdating = False
    Application.Calculate
    
    Dim seasonName As String
    Dim numberOfSection As Integer
    Dim pictureRangeSchedule, pictureRangeRanking As ChartObject
    Dim pictureName As String
    Dim minFileSize As Long
    
    seasonName = ActiveSheet.Cells(1, "A").Value
    numberOfSection = WorksheetFunction.CountIf(ActiveSheet.Range("BA2:BA241"), 0) / 8
    
    pictureName = "C:\Users\TaiNo\�}�C�h���C�u\�������A�^�C��_���M�ҋ@\schedule.jpg"
    If Dir(pictureName) <> "" Then
        MsgBox "��O���������܂����i3001�j"
        End
    End If
    
    Range("A" & WorksheetFunction.Max(1, numberOfSection * 8 - 6) & ":AG" & WorksheetFunction.Max(41, numberOfSection * 8 - 6 + 57)).CopyPicture '39
    Set pictureRangeSchedule = Sheets("�A�N�V�f���g").ChartObjects.Add(0, 0, Range("A" & WorksheetFunction.Max(1, numberOfSection * 8 - 6) & ":AG" & WorksheetFunction.Max(41, numberOfSection * 8 - 6 + 39)).Width, Range("A" & WorksheetFunction.Max(1, numberOfSection * 8 - 6) & ":AG" & WorksheetFunction.Max(41, numberOfSection * 8 - 6 + 57)).Height)
    pictureRangeSchedule.Chart.Export pictureName
    minFileSize = FileLen(pictureName)
    
    Do Until FileLen(pictureName) > minFileSize
        pictureRangeSchedule.Chart.Paste
        pictureRangeSchedule.Chart.Export pictureName
        DoEvents
    Loop
    
    pictureRangeSchedule.Delete
    Set pictureRangeSchedule = Nothing
    
    pictureName = "C:\Users\TaiNo\�}�C�h���C�u\�������A�^�C��_���M�ҋ@\ranking.jpg"
    If Dir(pictureName) <> "" Then
        MsgBox "��O���������܂����i3002�j"
        End
    End If
    
    Sheets(seasonName & "_�e��L�^").Range("A1:AR41").CopyPicture
    Set pictureRangeRanking = Sheets("�A�N�V�f���g").ChartObjects.Add(0, 0, Sheets(seasonName & "_�e��L�^").Range("A1:AR41").Width, Sheets(seasonName & "_�e��L�^").Range("A1:AR41").Height)
    pictureRangeRanking.Chart.Export pictureName
    minFileSize = FileLen(pictureName)
    
    Do Until FileLen(pictureName) > minFileSize
        pictureRangeRanking.Chart.Paste
        pictureRangeRanking.Chart.Export pictureName
        DoEvents
    Loop
    
    pictureRangeRanking.Delete
    Set pictureRangeRanking = Nothing
    
    Open "C:\Users\TaiNo\�}�C�h���C�u\�������A�^�C��_���M�ҋ@\nextGame.txt" For Output As #2
        Print #2, "�y�R�~�b�V���i�[���z"
        Print #2, "���������̒����ɂ����͂����肢���܂��B"
        Print #2, ""
        Print #2, "[��" & numberOfSection + 1 & "��]"
        If ActiveSheet.Cells(8 * numberOfSection + 3, "F") <> "" Then
            Print #2, "<���{��>�@" & ActiveSheet.Cells(8 * numberOfSection + 2, "C") & " " & ActiveSheet.Cells(8 * numberOfSection + 3, "D") & " - " & ActiveSheet.Cells(8 * numberOfSection + 3, "H") & " " & ActiveSheet.Cells(8 * numberOfSection + 2, "J")
        Else
            Print #2, ActiveSheet.Cells(8 * numberOfSection + 2, "C") & "(" & ActiveSheet.Cells(8 * numberOfSection + 2, "D") & ") - (" & ActiveSheet.Cells(8 * numberOfSection + 2, "H") & ")" & ActiveSheet.Cells(8 * numberOfSection + 2, "J")
        End If
        If ActiveSheet.Cells(8 * numberOfSection + 7, "F") <> "" Then
            Print #2, "<���{��>�@" & ActiveSheet.Cells(8 * numberOfSection + 6, "C") & " " & ActiveSheet.Cells(8 * numberOfSection + 7, "D") & " - " & ActiveSheet.Cells(8 * numberOfSection + 7, "H") & " " & ActiveSheet.Cells(8 * numberOfSection + 6, "J")
        Else
            Print #2, ActiveSheet.Cells(8 * numberOfSection + 6, "C") & "(" & ActiveSheet.Cells(8 * numberOfSection + 6, "D") & ") - (" & ActiveSheet.Cells(8 * numberOfSection + 6, "H") & ")" & ActiveSheet.Cells(8 * numberOfSection + 6, "J")
        End If
        Print #2, ""
        If numberOfSection < 29 Then
            Print #2, "[��" & numberOfSection + 2 & "��]"
            Print #2, ActiveSheet.Cells(8 * numberOfSection + 10, "C") & "(" & ActiveSheet.Cells(8 * numberOfSection + 10, "D") & ") - (" & ActiveSheet.Cells(8 * numberOfSection + 10, "H") & ")" & ActiveSheet.Cells(8 * numberOfSection + 10, "J")
            Print #2, ActiveSheet.Cells(8 * numberOfSection + 14, "C") & "(" & ActiveSheet.Cells(8 * numberOfSection + 14, "D") & ") - (" & ActiveSheet.Cells(8 * numberOfSection + 14, "H") & ")" & ActiveSheet.Cells(8 * numberOfSection + 14, "J");
        End If
    Close #2
    
    Call �o�b�N�A�b�v
    
    Application.ScreenUpdating = True

End Sub

