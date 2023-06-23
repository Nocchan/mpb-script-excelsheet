Attribute VB_Name = "mpb_vbascript_matchCompletion"
Dim season As String
Dim game As Integer
Dim section As Integer

Dim MPB_WORK_DIRECTORY_PATH As String
Dim DICT_TEAMNAME As New Dictionary

Dim BASE_ACCIDENT_RATE As Long
Dim DICT_ACCIDENT_HDCP As New Dictionary
Dim DICT_ACCIDENT_COEFFICIENT As New Dictionary

Dim DICT_ACCIDENT_LENGTH_RATE As New Dictionary
Dim DICT_ACCIDENT_MARGIN_DICT As New Dictionary
Dim DICT_ACCIDENT_MARGIN_S As New Dictionary
Dim DICT_ACCIDENT_MARGIN_A As New Dictionary
Dim DICT_ACCIDENT_MARGIN_B As New Dictionary
Dim DICT_ACCIDENT_MARGIN_C As New Dictionary
Dim DICT_ACCIDENT_MARGIN_D As New Dictionary
Dim DICT_ACCIDENT_MARGIN_E As New Dictionary
Dim DICT_ACCIDENT_MARGIN_F As New Dictionary
Dim DICT_ACCIDENT_MARGIN_G As New Dictionary
Dim DICT_ACCIDENT_MARGIN_n As New Dictionary

Sub matchCompletion()
    
    ' �f�o�b�O���[�h
    ' Call DebugMode

    ' �ďo���m�F
    If Not IsScheduleSheet() Then
        Call MessageError("�ďo���m�F�G���[", "matchCompletion")
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
    
    MPB_WORK_DIRECTORY_PATH = "C:\Users\TaiNo\�}�C�h���C�u\MPB\1-�܂�"
    
    With DICT_TEAMNAME
        .Add "G", "�W���C�A���c"
        .Add "M", "�}���[���Y"
        .Add "T", "�^�C�K�[�X"
        .Add "L", "���C�I���Y"
        .Add "E", "�C�[�O���X"
    End With
    
    BASE_ACCIDENT_RATE = 0.01
    
    With DICT_ACCIDENT_HDCP
        .Add "G", 1#
        .Add "M", 1#
        .Add "T", 1#
        .Add "L", 1#
        .Add "E", 1#
    End With
    
    With DICT_ACCIDENT_COEFFICIENT
        .Add "S", 0.01
        .Add "A", 0.3
        .Add "B", 0.5
        .Add "C", 0.8
        .Add "D", 1#
        .Add "E", 1.2
        .Add "F", 2#
        .Add "G", 4#
        .Add "n", 0#
    End With
    
    With DICT_ACCIDENT_LENGTH_RATE
        .Add 1, 30#
        .Add 2, 49#
        .Add 5, 10.5
        .Add 8, 7
        .Add 24, 3.5
    End With
    
    With DICT_ACCIDENT_MARGIN_DICT
        .Add "S", DICT_ACCIDENT_MARGIN_S
        .Add "A", DICT_ACCIDENT_MARGIN_A
        .Add "B", DICT_ACCIDENT_MARGIN_B
        .Add "C", DICT_ACCIDENT_MARGIN_C
        .Add "D", DICT_ACCIDENT_MARGIN_D
        .Add "E", DICT_ACCIDENT_MARGIN_E
        .Add "F", DICT_ACCIDENT_MARGIN_F
        .Add "G", DICT_ACCIDENT_MARGIN_G
        .Add "n", DICT_ACCIDENT_MARGIN_n
    End With
    
    With DICT_ACCIDENT_MARGIN_S
        .Add -1, 30#
        .Add 0, 70#
    End With
    
    With DICT_ACCIDENT_MARGIN_A
        .Add -1, 30#
        .Add 0, 40#
        .Add 1, 30#
    End With
    
    With DICT_ACCIDENT_MARGIN_B
        .Add -1, 30#
        .Add 0, 40#
        .Add 1, 20#
        .Add 2, 10#
    End With
    
    With DICT_ACCIDENT_MARGIN_C
        .Add -1, 30#
        .Add 0, 40#
        .Add 1, 15#
        .Add 2, 10#
        .Add 3, 5#
    End With
    
    With DICT_ACCIDENT_MARGIN_D
        .Add -1, 30#
        .Add 0, 40#
        .Add 1, 12#
        .Add 2, 9#
        .Add 3, 6#
        .Add 4, 3#
    End With
    
    With DICT_ACCIDENT_MARGIN_E
        .Add -1, 30#
        .Add 0, 40#
        .Add 1, 10#
        .Add 2, 8#
        .Add 3, 6#
        .Add 4, 4#
        .Add 5, 2#
    End With
    
    With DICT_ACCIDENT_MARGIN_F
        .Add -1, 30#
        .Add 0, 40#
        .Add 1, 8.57
        .Add 2, 7.14
        .Add 3, 5.71
        .Add 4, 4.29
        .Add 5, 2.86
        .Add 6, 1.43
    End With
    
    With DICT_ACCIDENT_MARGIN_G
        .Add -1, 30#
        .Add 0, 40#
        .Add 1, 7.5
        .Add 2, 6.43
        .Add 3, 5.36
        .Add 4, 4.29
        .Add 5, 3.21
        .Add 6, 2.14
        .Add 7, 1.07
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
        Call MessageError("�s�����̓G���[", "IsSectionCompleted")
        Call ExitProcess
    End If
    
    ' �J���O�܂��͍ŏI�ߌ�ŗ\���攭���l����K�v���Ȃ��p�^�[��
    If section = 0 Or section = 30 Then
        IsSectionCompleted = True
        Exit Function
    End If
    
    ' �\���攭���o�����Ă��Ȃ��p�^�[��
    If ActiveSheet.Cells(section * 8 + 2, "D").Value = "" Or ActiveSheet.Cells(section * 8 + 6, "D").Value = "" Or _
       ActiveSheet.Cells(section * 8 + 2, "H").Value = "" Or ActiveSheet.Cells(section * 8 + 6, "H").Value = "" Then
        Call MessageError("�\���攭�������G���[", "IsSectionCompleted")
        Call ExitProcess
    End If
    
    IsSectionCompleted = True
    
End Function

' �߂̐i�s�ɂ�蔭������A���炩���ߗ\�肳��Ă���C�x���g���o��
Function MakeMPBNewsSeasonEvent()
    
    Dim mpbNewsSeasonEventFlg As Boolean
    Dim mpbNewsSeasonEvent As String
    
    mpbNewsSeasonEventFlg = False
    mpbNewsSeasonEvent = "�yMPB�^�c����̂��m�点�z"

    If section = 10 Or section = 20 Then
        mpbNewsSeasonEventFlg = True
        mpbNewsSeasonEvent = AddRowText(mpbNewsSeasonEvent, "�ETSOB�g�̐U�蒼�����s���܂��BTSOB�g�̕\���ݒ���ŐV�����Ă��������B")
        mpbNewsSeasonEvent = AddRowText(mpbNewsSeasonEvent, "- - - - - - - - - -")
        mpbNewsSeasonEvent = AddRowText(mpbNewsSeasonEvent, "1��: " & Left(Sheets(season & "_�e��L�^").Cells(2, "B").Value, 1) & " �� 3.5")
        mpbNewsSeasonEvent = AddRowText(mpbNewsSeasonEvent, "2��: " & Left(Sheets(season & "_�e��L�^").Cells(3, "B").Value, 1) & " �� 4.0")
        mpbNewsSeasonEvent = AddRowText(mpbNewsSeasonEvent, "3��: " & Left(Sheets(season & "_�e��L�^").Cells(4, "B").Value, 1) & " �� 4.5")
        mpbNewsSeasonEvent = AddRowText(mpbNewsSeasonEvent, "4��: " & Left(Sheets(season & "_�e��L�^").Cells(5, "B").Value, 1) & " �� 5.0")
        mpbNewsSeasonEvent = AddRowText(mpbNewsSeasonEvent, "5��: " & Left(Sheets(season & "_�e��L�^").Cells(6, "B").Value, 1) & " �� 5.5")
        mpbNewsSeasonEvent = AddRowText(mpbNewsSeasonEvent, "�������`�[���������ɂ́A�K���������̒ʂ�ƂȂ�Ȃ��ꍇ������܂��B")
        mpbNewsSeasonEvent = AddRowText(mpbNewsSeasonEvent, "")
    End If
    
    If section = 10 Or section = 20 Then
        mpbNewsSeasonEventFlg = True
        mpbNewsSeasonEvent = AddRowText(mpbNewsSeasonEvent, "�E�������A�㔼�킩���HDCP�ύX��t���J�n���܂��B��15�ߏI���������Ē��ߐ؂�̂ŁA�ύX�������`�[���́A�K�v�ɉ����Đ\�����s���Ă��������B�ύX���Ȃ��ꍇ�́A���ɑΉ��s�v�ł��B")
        mpbNewsSeasonEvent = AddRowText(mpbNewsSeasonEvent, "")
    End If
    
    If section = 15 Then
        mpbNewsSeasonEventFlg = True
        mpbNewsSeasonEvent = AddRowText(mpbNewsSeasonEvent, "�E�����������܂��āA�㔼��Ɍ�����HDCP�ύX�̐\������ߐ؂�܂��BHDCP�̕\���ݒ���ŐV�����Ă��������B")
        mpbNewsSeasonEvent = AddRowText(mpbNewsSeasonEvent, "")
    End If
    
    If section = 25 Then
        mpbNewsSeasonEventFlg = True
        mpbNewsSeasonEvent = AddRowText(mpbNewsSeasonEvent, "�E�������AB9GG�m�~�l�[�g�I�[�_�[�̒�o��t���J�n���܂��B��28�ߏI���������Ē��ߐ؂�̂ŁA�e�`�[���ALINE�O���[�v�̃A���o���u" & season & "B9GG�m�~�l�[�g�v�ɒ�o�����肢�������܂��B")
        mpbNewsSeasonEvent = AddRowText(mpbNewsSeasonEvent, "")
    End If
    
    If section = 26 Or section = 27 Then
        mpbNewsSeasonEventFlg = True
        mpbNewsSeasonEvent = AddRowText(mpbNewsSeasonEvent, "�EB9GG�m�~�l�[�g�I�[�_�[�̒�o/�ύX����t���ł��B����o�̃`�[���́A��28�߂��I������܂łɁALINE�O���[�v�̃A���o���u" & season & "B9GG�m�~�l�[�g�v�ւ̒�o�����肢�������܂��B")
        mpbNewsSeasonEvent = AddRowText(mpbNewsSeasonEvent, "")
    End If
    
    If section = 28 Then
        mpbNewsSeasonEventFlg = True
        mpbNewsSeasonEvent = AddRowText(mpbNewsSeasonEvent, "�EB9GG�m�~�l�[�g�I�[�_�[���o��t���ł��B����o�̃`�[���́A��28�߂��I������܂łɁALINE�O���[�v�̃A���o���u" & season & "B9GG�m�~�l�[�g�v�ւ̒�o�����肢�������܂��B")
        mpbNewsSeasonEvent = AddRowText(mpbNewsSeasonEvent, "")
    End If
    
    If section = 30 Then
        mpbNewsSeasonEventFlg = True
        mpbNewsSeasonEvent = AddRowText(mpbNewsSeasonEvent, "�E���V�[�Y���A�\�肳��Ă����S�������I�����܂����B�܂��́A�F���񂨔�ꂳ�܂ł����I���̌�AMPB�A���[�h�����{���܂��̂ŁA�ē������҂����������B")
        mpbNewsSeasonEvent = AddRowText(mpbNewsSeasonEvent, "")
    End If
    
    If mpbNewsSeasonEventFlg Then
        mpbNewsSeasonEvent = AddRowText(mpbNewsSeasonEvent, "�ȏ�")
        Call OutputText(mpbNewsSeasonEvent, MPB_WORK_DIRECTORY_PATH & "\mpbnews-seasonevent.txt")
    End If

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

