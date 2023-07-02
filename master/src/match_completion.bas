Attribute VB_Name = "match_completion"
Option Explicit

Dim season As String
Dim game As Integer
Dim section As Integer

Dim dictTeamID As New Dictionary

Sub MatchCompletion()

    ' �f�o�b�O���[�h
    Call enableDebugMode

    ' �ďo���m�F
    If Not isScheduleSheet() Then
        Call showMessageError("�ďo���m�F�G���[", "MatchCompletion")
        End
    End If

    Call initialize

    If isSectionCompleted() Then
        Call makeMPBNewsSeasonEvent
        Call makeMPBNewsOfThisSection
        Call makeMPBNewsOfAccident
    End If

    Call makeMPBNewsOfNextGame
    Call savePictureOfSchedule
    Call savePictureOfRecord

    Call exitProcess

End Sub

' �萔�E�V�[�g��Ԃ̏�����
Function initialize()

    If Not isDebugMode Then
        Application.ScreenUpdating = False
    End If

    Application.Calculate

    Call makeBackupFile

    season = ActiveSheet.Cells(1, "A").Value
    game = WorksheetFunction.CountIf(ActiveSheet.Range("F2:F241"), "�i�����I���j")
    section = WorksheetFunction.CountIf(ActiveSheet.Range("BA2:BA241"), 0) / 8

    Dim teamID As Integer
    For teamID = 1 To 5
        dictTeamID.Add teamID, Sheets(season & "_�e��L�^").Cells(teamID + 1, "R").Value
    Next teamID

    Call Definition

    Sheets(season & "_�X�P�W���[��").Unprotect
    Sheets(season & "_����f�[�^").Unprotect
    Sheets(season & "_���f�[�^").Unprotect

End Function

' �I��������
Function exitProcess()

    Sheets(season & "_�X�P�W���[��").Protect AllowFormattingColumns:=True, AllowFormattingRows:=True
    Sheets(season & "_����f�[�^").Protect AllowFormattingColumns:=True, AllowFormattingRows:=True
    Sheets(season & "_���f�[�^").Protect AllowFormattingColumns:=True, AllowFormattingRows:=True

    If Not isDebugMode Then
        Application.ScreenUpdating = True
    End If

    Application.Calculate

    End

End Function

' �߂��������ăX�y������s�����Ԃ��𔻒�
Function isSectionCompleted() As Boolean

    ' ��������������߂��������Ă��Ȃ����Ƃ��킩��p�^�[��
    If game <> section * 2 Then
        isSectionCompleted = False
        Exit Function
    End If

    ' �߂͊������Ă��邪�s�����͂�����p�^�[��
    If Sheets(season & "_�X�P�W���[��").Cells(section * 8 + 3, "D").Value <> "" Or Sheets(season & "_�X�P�W���[��").Cells(section * 8 + 7, "D").Value <> "" Or _
       Sheets(season & "_�X�P�W���[��").Cells(section * 8 + 3, "F").Value <> "" Or Sheets(season & "_�X�P�W���[��").Cells(section * 8 + 7, "F").Value <> "" Or _
       Sheets(season & "_�X�P�W���[��").Cells(section * 8 + 3, "H").Value <> "" Or Sheets(season & "_�X�P�W���[��").Cells(section * 8 + 7, "H").Value <> "" Then
        Call showMessageError("�s�����̓G���[", "isSectionCompleted")
        Call exitProcess
    End If

    ' �J���O�܂��͍ŏI�ߌ�ŗ\���攭���l����K�v���Ȃ��p�^�[��
    If section = 0 Or section = 30 Then
        isSectionCompleted = True
        Exit Function
    End If

    ' �\���攭���o�����Ă��Ȃ��p�^�[��
    If Sheets(season & "_�X�P�W���[��").Cells(section * 8 + 2, "D").Value = "" Or Sheets(season & "_�X�P�W���[��").Cells(section * 8 + 6, "D").Value = "" Or _
       Sheets(season & "_�X�P�W���[��").Cells(section * 8 + 2, "H").Value = "" Or Sheets(season & "_�X�P�W���[��").Cells(section * 8 + 6, "H").Value = "" Then
        Call showMessageError("�\���攭�������G���[", "isSectionCompleted")
        Call exitProcess
    End If

    isSectionCompleted = True

End Function

' �߂̐i�s�ɂ�蔭������A���炩���ߗ\�肳��Ă���C�x���g���o��
Function makeMPBNewsSeasonEvent()

    ' �錾
    Dim existMPBNewsSeasonEvent As Boolean
    Dim bodyMPBNewsSeasonEvent As String
    Dim tsobBorderDict As New Dictionary

    ' ������
    existMPBNewsSeasonEvent = False
    bodyMPBNewsSeasonEvent = "�yMPB�^�c����̂��m�点�z"

    ' TSOB�g�̐U�蒼��
    If section = 10 Or section = 20 Then
        existMPBNewsSeasonEvent = True

        bodyMPBNewsSeasonEvent = addLineToText(bodyMPBNewsSeasonEvent, "")
        bodyMPBNewsSeasonEvent = addLineToText(bodyMPBNewsSeasonEvent, "�ETSOB�g�̐U�蒼�����s���܂��B")
        bodyMPBNewsSeasonEvent = addLineToText(bodyMPBNewsSeasonEvent, "1��: " & Left(Sheets(season & "_�e��L�^").Cells(2, "B").Value, 1) & " �� 3.5")
        bodyMPBNewsSeasonEvent = addLineToText(bodyMPBNewsSeasonEvent, "2��: " & Left(Sheets(season & "_�e��L�^").Cells(3, "B").Value, 1) & " �� 4.0")
        bodyMPBNewsSeasonEvent = addLineToText(bodyMPBNewsSeasonEvent, "3��: " & Left(Sheets(season & "_�e��L�^").Cells(4, "B").Value, 1) & " �� 4.5")
        bodyMPBNewsSeasonEvent = addLineToText(bodyMPBNewsSeasonEvent, "4��: " & Left(Sheets(season & "_�e��L�^").Cells(5, "B").Value, 1) & " �� 5.0")
        bodyMPBNewsSeasonEvent = addLineToText(bodyMPBNewsSeasonEvent, "5��: " & Left(Sheets(season & "_�e��L�^").Cells(6, "B").Value, 1) & " �� 5.5")
        bodyMPBNewsSeasonEvent = addLineToText(bodyMPBNewsSeasonEvent, "�������`�[���������ɂ́A�K���������̒ʂ�ƂȂ�Ȃ��ꍇ������܂��B")

        tsobBorderDict.Add Left(Sheets(season & "_�e��L�^").Cells(2, "B").Value, 1), "3.5"
        tsobBorderDict.Add Left(Sheets(season & "_�e��L�^").Cells(3, "B").Value, 1), "4.0"
        tsobBorderDict.Add Left(Sheets(season & "_�e��L�^").Cells(4, "B").Value, 1), "4.5"
        tsobBorderDict.Add Left(Sheets(season & "_�e��L�^").Cells(5, "B").Value, 1), "5.0"
        tsobBorderDict.Add Left(Sheets(season & "_�e��L�^").Cells(6, "B").Value, 1), "5.5"

        Call showMessageDebug(tsobBorderDict.Item(Sheets(season & "_�X�P�W���[��").Cells(1, "BB").Value), "INPUT �X�P�W���[��.Cells(27,CP)")
        Sheets(season & "_�X�P�W���[��").Cells(27, "CP").Value = tsobBorderDict.Item(Sheets(season & "_�X�P�W���[��").Cells(1, "BB").Value)
        Call showMessageDebug(tsobBorderDict.Item(Sheets(season & "_�X�P�W���[��").Cells(1, "BC").Value), "INPUT �X�P�W���[��.Cells(27,CQ)")
        Sheets(season & "_�X�P�W���[��").Cells(27, "CQ").Value = tsobBorderDict.Item(Sheets(season & "_�X�P�W���[��").Cells(1, "BC").Value)
        Call showMessageDebug(tsobBorderDict.Item(Sheets(season & "_�X�P�W���[��").Cells(1, "BD").Value), "INPUT �X�P�W���[��.Cells(27,CR)")
        Sheets(season & "_�X�P�W���[��").Cells(27, "CR").Value = tsobBorderDict.Item(Sheets(season & "_�X�P�W���[��").Cells(1, "BD").Value)
        Call showMessageDebug(tsobBorderDict.Item(Sheets(season & "_�X�P�W���[��").Cells(1, "BE").Value), "INPUT �X�P�W���[��.Cells(27,CS)")
        Sheets(season & "_�X�P�W���[��").Cells(27, "CS").Value = tsobBorderDict.Item(Sheets(season & "_�X�P�W���[��").Cells(1, "BE").Value)
        Call showMessageDebug(tsobBorderDict.Item(Sheets(season & "_�X�P�W���[��").Cells(1, "BF").Value), "INPUT �X�P�W���[��.Cells(27,CT)")
        Sheets(season & "_�X�P�W���[��").Cells(27, "CT").Value = tsobBorderDict.Item(Sheets(season & "_�X�P�W���[��").Cells(1, "BF").Value)
    End If

    ' HDCP�ύX��t�J�n
    If section = 10 Or section = 20 Then
        existMPBNewsSeasonEvent = True

        bodyMPBNewsSeasonEvent = addLineToText(bodyMPBNewsSeasonEvent, "")
        bodyMPBNewsSeasonEvent = addLineToText(bodyMPBNewsSeasonEvent, "�E�������A�㔼�킩���HDCP�ύX��t���J�n���܂��B��15�ߏI���������Ē��ߐ؂�̂ŁA�ύX�������`�[���́A�K�v�ɉ����Đ\�����s���Ă��������B�ύX���Ȃ��ꍇ�́A���ɑΉ��s�v�ł��B")
    End If

    ' HDCP�ύX��
    If section = 11 Or section = 12 Or section = 13 Or section = 14 Then
        existMPBNewsSeasonEvent = True

        bodyMPBNewsSeasonEvent = addLineToText(bodyMPBNewsSeasonEvent, "")
        bodyMPBNewsSeasonEvent = addLineToText(bodyMPBNewsSeasonEvent, "�E�㔼�킩���HDCP�ύX����t���ł��B�ύX�������`�[���́A��15�ߏI���܂łɐ\�����s���Ă��������B")
    End If

    ' HDCP�ύX��t�Y
    If section = 15 Then
        existMPBNewsSeasonEvent = True

        bodyMPBNewsSeasonEvent = addLineToText(bodyMPBNewsSeasonEvent, "")
        bodyMPBNewsSeasonEvent = addLineToText(bodyMPBNewsSeasonEvent, "�E�����������܂��āA�㔼��Ɍ�����HDCP�ύX�̐\������ߐ؂�܂��BHDCP�̕\���ݒ���ŐV�����Ă��������B")
    End If

    ' B9GG��o��t�J�n
    If section = 25 Then
        existMPBNewsSeasonEvent = True

        bodyMPBNewsSeasonEvent = addLineToText(bodyMPBNewsSeasonEvent, "")
        bodyMPBNewsSeasonEvent = addLineToText(bodyMPBNewsSeasonEvent, "�E�������AB9GG�m�~�l�[�g�I�[�_�[�̒�o��t���J�n���܂��B��28�ߏI���������Ē��ߐ؂�̂ŁA�e�`�[���ALINE�O���[�v�̃A���o���u" & season & "B9GG�m�~�l�[�g�v�ɒ�o�����肢�������܂��B")
    End If

    ' B9GG��o��t��
    If section = 26 Or section = 27 Then
        existMPBNewsSeasonEvent = True

        bodyMPBNewsSeasonEvent = addLineToText(bodyMPBNewsSeasonEvent, "")
        bodyMPBNewsSeasonEvent = addLineToText(bodyMPBNewsSeasonEvent, "�EB9GG�m�~�l�[�g�I�[�_�[�̒�o/�ύX����t���ł��B����o�̃`�[���́A��28�߂��I������܂łɁALINE�O���[�v�̃A���o���u" & season & "B9GG�m�~�l�[�g�v�ւ̒�o�����肢�������܂��B")
    End If

    ' B9GG��o��t�Y
    If section = 28 Then
        existMPBNewsSeasonEvent = True

        bodyMPBNewsSeasonEvent = addLineToText(bodyMPBNewsSeasonEvent, "")
        bodyMPBNewsSeasonEvent = addLineToText(bodyMPBNewsSeasonEvent, "�E�����������܂��āAB9GG�m�~�l�[�g�I�[�_�[�̒�o����ߐ؂�܂��B")
    End If

    ' MPB�A���[�h�ē�
    If section = 30 Then
        existMPBNewsSeasonEvent = True

        bodyMPBNewsSeasonEvent = addLineToText(bodyMPBNewsSeasonEvent, "")
        bodyMPBNewsSeasonEvent = addLineToText(bodyMPBNewsSeasonEvent, "�E���V�[�Y���A�\�肳��Ă����S�������I�����܂����B�܂��́A�F���񂨔�ꂳ�܂ł����I���̌�AMPB�A���[�h�����{���܂��̂ŁA�ē������҂����������B")
    End If

    ' ���ʂ̏o��
    If existMPBNewsSeasonEvent Then
        bodyMPBNewsSeasonEvent = addLineToText(bodyMPBNewsSeasonEvent, "")
        bodyMPBNewsSeasonEvent = addLineToText(bodyMPBNewsSeasonEvent, "�ȏ�")

        If Not isDebugMode Then
            Call saveTxtFile(bodyMPBNewsSeasonEvent, MPB_WORK_DIRECTORY_PATH & "\batch-min\mpbnews-seasonevent.txt")
        Else
            Call showMessageInfo(bodyMPBNewsSeasonEvent, "makeMPBNewsSeasonEvent")
            Call saveTxtFile(bodyMPBNewsSeasonEvent, LOCAL_WORK_DIRECTORY_PATH & "\batch-min\mpbnews-seasonevent.txt")
        End If
    End If

End Function

' �߂̐i�s�ɂ�蔭������A�D���}�W�b�N�⎩�͗D���Ɋւ���C�x���g���o��
Function makeMPBNewsOfThisSection()

    ' ���s����
    If section = 0 Then
        Exit Function
    End If

    ' �錾
    Dim bodyMPBNewsOfThisSection As String
    Dim scoreOfThisSection(2, 4) As String
    Dim teamOfNextSection(2, 2) As String
    Dim seasonStatus As New Dictionary
    Dim headerOfNextGame As String
    Dim messageTemplateVictory As String
    Dim messageTemplateMagicDisappearance As String
    Dim messageTemplateSelfVictoryDisappearance As String
    Dim messageTemplateSelfVictoryReappearance As String
    Dim messageTemplateMagicAppearance As String

    ' ������
    bodyMPBNewsOfThisSection = "�yMPB�j���[�X�z"
    messageTemplateVictory = season & "�y�i���g���[�X�D�����m��I"
    messageTemplateMagicDisappearance = "�D���}�W�b�N�����Łc"
    messageTemplateSelfVictoryDisappearance = "���͗D�������Łc"
    messageTemplateSelfVictoryReappearance = "���͗D���������I"
    messageTemplateMagicAppearance = "�D���}�W�b�N���_���I"
    
    ' ���߂̎�������
    If section > 0 Then
        scoreOfThisSection(1, 1) = Sheets(season & "_�X�P�W���[��").Cells(section * 8 - 6, "C").Value
        scoreOfThisSection(1, 2) = Sheets(season & "_�X�P�W���[��").Cells(section * 8 - 5, "D").Value
        scoreOfThisSection(1, 3) = Sheets(season & "_�X�P�W���[��").Cells(section * 8 - 5, "H").Value
        scoreOfThisSection(1, 4) = Sheets(season & "_�X�P�W���[��").Cells(section * 8 - 6, "J").Value
        scoreOfThisSection(2, 1) = Sheets(season & "_�X�P�W���[��").Cells(section * 8 - 2, "C").Value
        scoreOfThisSection(2, 2) = Sheets(season & "_�X�P�W���[��").Cells(section * 8 - 1, "D").Value
        scoreOfThisSection(2, 3) = Sheets(season & "_�X�P�W���[��").Cells(section * 8 - 1, "H").Value
        scoreOfThisSection(2, 4) = Sheets(season & "_�X�P�W���[��").Cells(section * 8 - 2, "J").Value
        bodyMPBNewsOfThisSection = addLineToText(bodyMPBNewsOfThisSection, "")
        bodyMPBNewsOfThisSection = addLineToText(bodyMPBNewsOfThisSection, "��" & section & "�߂̎�������")
        bodyMPBNewsOfThisSection = addLineToText(bodyMPBNewsOfThisSection, scoreOfThisSection(1, 1) & " " & scoreOfThisSection(1, 2) & "-" & scoreOfThisSection(1, 3) & " " & scoreOfThisSection(1, 4))
        bodyMPBNewsOfThisSection = addLineToText(bodyMPBNewsOfThisSection, scoreOfThisSection(2, 1) & " " & scoreOfThisSection(2, 2) & "-" & scoreOfThisSection(2, 3) & " " & scoreOfThisSection(2, 4))
    End If

    ' ���߂̗\���攭
    If section < 30 Then
        teamOfNextSection(1, 1) = Sheets(season & "_�X�P�W���[��").Cells(section * 8 + 2, "C").Value
        teamOfNextSection(1, 2) = Sheets(season & "_�X�P�W���[��").Cells(section * 8 + 2, "J").Value
        teamOfNextSection(2, 1) = Sheets(season & "_�X�P�W���[��").Cells(section * 8 + 6, "C").Value
        teamOfNextSection(2, 2) = Sheets(season & "_�X�P�W���[��").Cells(section * 8 + 6, "J").Value
        bodyMPBNewsOfThisSection = addLineToText(bodyMPBNewsOfThisSection, "")
        bodyMPBNewsOfThisSection = addLineToText(bodyMPBNewsOfThisSection, "��" & section + 1 & "�߂̗\���攭")
        bodyMPBNewsOfThisSection = addLineToText(bodyMPBNewsOfThisSection, "�i" & Sheets(season & "_�X�P�W���[��").Cells(section * 8 + 2, "C").Value & "-" & Sheets(season & "_�X�P�W���[��").Cells(section * 8 + 2, "J").Value & "�j" & Sheets(season & "_�X�P�W���[��").Cells(section * 8 + 2, "D").Value & "�~" & Sheets(season & "_�X�P�W���[��").Cells(section * 8 + 2, "H").Value)
        bodyMPBNewsOfThisSection = addLineToText(bodyMPBNewsOfThisSection, "�i" & Sheets(season & "_�X�P�W���[��").Cells(section * 8 + 6, "C").Value & "-" & Sheets(season & "_�X�P�W���[��").Cells(section * 8 + 6, "J").Value & "�j" & Sheets(season & "_�X�P�W���[��").Cells(section * 8 + 6, "D").Value & "�~" & Sheets(season & "_�X�P�W���[��").Cells(section * 8 + 6, "H").Value)
        bodyMPBNewsOfThisSection = addLineToText(bodyMPBNewsOfThisSection, "")
    End If

    ' �󋵊m�F(���ߎ��{�O)
    seasonStatus.Add "���ߎ��{�O", seasonStatusOfSection(section, "", "", "", "", "", "")

    ' ���ߎ��{�O�ɗD�������܂��Ă��Ȃ��O��
    If seasonStatus.Item("���ߎ��{�O")(0) = "" Then

        seasonStatus.Add "���ߎ��{��", seasonStatusOfSection(section, scoreOfThisSection(1, 2), Sheets(season & "_�X�P�W���[��").Cells(section * 8 - 5, "F").Value, scoreOfThisSection(1, 3), scoreOfThisSection(2, 2), Sheets(season & "_�X�P�W���[��").Cells(section * 8 - 1, "F").Value, scoreOfThisSection(2, 3))
    
        ' ���߂ŗD�������܂����ꍇ
        If seasonStatus.Item("���ߎ��{��")(0) <> "" Then
            bodyMPBNewsOfThisSection = addLineToText(bodyMPBNewsOfThisSection, "��" & seasonStatus.Item("���ߎ��{��")(0) & "��" & season & "�y�i���g���[�X�D�����m��I")
        Else
            Dim teamID As Integer
            seasonStatus.Add "����AA", seasonStatusOfSection(section + 1, "X", "tmp", "0", "X", "tmp", "0")
            seasonStatus.Add "����BA", seasonStatusOfSection(section + 1, "0", "tmp", "X", "X", "tmp", "0")
            seasonStatus.Add "����AB", seasonStatusOfSection(section + 1, "X", "tmp", "0", "0", "tmp", "X")
            seasonStatus.Add "����BB", seasonStatusOfSection(section + 1, "0", "tmp", "X", "0", "tmp", "X")
            
            ' ����
            headerOfNextGame = ""
            ' �}�W�b�N�����ł���P�[�X
            For teamID = 1 To 5
                If Left(seasonStatus.Item("���ߎ��{�O")(teamID), 1) = "M" And Left(seasonStatus.Item("���ߎ��{��")(teamID), 1) <> "M" Then
                    bodyMPBNewsOfThisSection = addLineToText(bodyMPBNewsOfThisSection, "��" & DICT_TEAM_NAME.Item(dictTeamID.Item(teamID)) & "��" & headerOfNextGame & messageTemplateMagicDisappearance)
                End If
            Next teamID
            
            ' ���͗D�������ł���P�[�X
            For teamID = 1 To 5
                If Left(seasonStatus.Item("���ߎ��{�O")(teamID), 1) <> "��" And Left(seasonStatus.Item("���ߎ��{��")(teamID), 1) = "��" Then
                    bodyMPBNewsOfThisSection = addLineToText(bodyMPBNewsOfThisSection, "��" & DICT_TEAM_NAME.Item(dictTeamID.Item(teamID)) & "��" & headerOfNextGame & messageTemplateSelfVictoryDisappearance)
                End If
            Next teamID
            
            ' ���͗D������������P�[�X
            For teamID = 1 To 5
                If Left(seasonStatus.Item("���ߎ��{�O")(teamID), 1) = "��" And Left(seasonStatus.Item("���ߎ��{��")(teamID), 1) <> "��" Then
                    bodyMPBNewsOfThisSection = addLineToText(bodyMPBNewsOfThisSection, "��" & DICT_TEAM_NAME.Item(dictTeamID.Item(teamID)) & "��" & headerOfNextGame & messageTemplateSelfVictoryReappearance)
                End If
            Next teamID
            
            ' �}�W�b�N���_������P�[�X
            For teamID = 1 To 5
                If Left(seasonStatus.Item("���ߎ��{�O")(teamID), 1) <> "M" And Left(seasonStatus.Item("���ߎ��{��")(teamID), 1) = "M" Then
                    bodyMPBNewsOfThisSection = addLineToText(bodyMPBNewsOfThisSection, "��" & DICT_TEAM_NAME.Item(dictTeamID.Item(teamID)) & "��" & headerOfNextGame & messageTemplateMagicAppearance)
                End If
            Next teamID
            
            ' ����A_
            headerOfNextGame = "���� " & teamOfNextSection(1, 1) & "��-��" & teamOfNextSection(1, 2) & " �ŁA"
            If seasonStatus.Item("����AA")(0) <> "" And seasonStatus.Item("����AB")(0) <> "" And seasonStatus.Item("����AA")(0) = seasonStatus.Item("����AB")(0) Then
                ' �D���`�[�������܂�P�[�X
                bodyMPBNewsOfThisSection = addLineToText(bodyMPBNewsOfThisSection, "��" & seasonStatus.Item("����AA")(0) & "��" & headerOfNextGame & messageTemplateVictory)
            Else
                ' �}�W�b�N�����ł���P�[�X
                For teamID = 1 To 5
                    If Left(seasonStatus.Item("���ߎ��{��")(teamID), 1) = "M" And Left(seasonStatus.Item("����AA")(teamID), 1) <> "M" And Left(seasonStatus.Item("����AB")(teamID), 1) <> "M" Then
                        bodyMPBNewsOfThisSection = addLineToText(bodyMPBNewsOfThisSection, "��" & DICT_TEAM_NAME.Item(dictTeamID.Item(teamID)) & "��" & headerOfNextGame & messageTemplateMagicDisappearance)
                    End If
                Next teamID
                
                ' ���͗D�������ł���P�[�X
                For teamID = 1 To 5
                    If Left(seasonStatus.Item("���ߎ��{��")(teamID), 1) <> "��" And Left(seasonStatus.Item("����AA")(teamID), 1) = "��" And Left(seasonStatus.Item("����AB")(teamID), 1) = "��" Then
                        bodyMPBNewsOfThisSection = addLineToText(bodyMPBNewsOfThisSection, "��" & DICT_TEAM_NAME.Item(dictTeamID.Item(teamID)) & "��" & headerOfNextGame & messageTemplateSelfVictoryDisappearance)
                    End If
                Next teamID
                
                ' ���͗D������������P�[�X
                For teamID = 1 To 5
                    If Left(seasonStatus.Item("���ߎ��{��")(teamID), 1) = "��" And Left(seasonStatus.Item("����AA")(teamID), 1) <> "��" And Left(seasonStatus.Item("����AB")(teamID), 1) <> "��" Then
                        bodyMPBNewsOfThisSection = addLineToText(bodyMPBNewsOfThisSection, "��" & DICT_TEAM_NAME.Item(dictTeamID.Item(teamID)) & "��" & headerOfNextGame & messageTemplateSelfVictoryReappearance)
                    End If
                Next teamID
                
                ' �}�W�b�N���_������P�[�X
                For teamID = 1 To 5
                    If Left(seasonStatus.Item("���ߎ��{��")(teamID), 1) <> "M" And Left(seasonStatus.Item("����AA")(teamID), 1) = "M" And Left(seasonStatus.Item("����AB")(teamID), 1) = "M" Then
                        bodyMPBNewsOfThisSection = addLineToText(bodyMPBNewsOfThisSection, "��" & DICT_TEAM_NAME.Item(dictTeamID.Item(teamID)) & "��" & headerOfNextGame & messageTemplateMagicAppearance)
                    End If
                Next teamID
            End If
            
            ' ����B_
            headerOfNextGame = "���� " & teamOfNextSection(1, 1) & "��-��" & teamOfNextSection(1, 2) & " �ŁA"
            If seasonStatus.Item("����BA")(0) <> "" And seasonStatus.Item("����BB")(0) <> "" And seasonStatus.Item("����BA")(0) = seasonStatus.Item("����BB")(0) Then
                ' �D���`�[�������܂�P�[�X
                bodyMPBNewsOfThisSection = addLineToText(bodyMPBNewsOfThisSection, "��" & seasonStatus.Item("����BB")(0) & "��" & headerOfNextGame & messageTemplateVictory)
            Else
                ' �}�W�b�N�����ł���P�[�X
                For teamID = 1 To 5
                    If Left(seasonStatus.Item("���ߎ��{��")(teamID), 1) = "M" And Left(seasonStatus.Item("����BA")(teamID), 1) <> "M" And Left(seasonStatus.Item("����BB")(teamID), 1) <> "M" Then
                        bodyMPBNewsOfThisSection = addLineToText(bodyMPBNewsOfThisSection, "��" & DICT_TEAM_NAME.Item(dictTeamID.Item(teamID)) & "��" & headerOfNextGame & messageTemplateMagicDisappearance)
                    End If
                Next teamID
                
                ' ���͗D�������ł���P�[�X
                For teamID = 1 To 5
                    If Left(seasonStatus.Item("���ߎ��{��")(teamID), 1) <> "��" And Left(seasonStatus.Item("����BA")(teamID), 1) = "��" And Left(seasonStatus.Item("����BB")(teamID), 1) = "��" Then
                        bodyMPBNewsOfThisSection = addLineToText(bodyMPBNewsOfThisSection, "��" & DICT_TEAM_NAME.Item(dictTeamID.Item(teamID)) & "��" & headerOfNextGame & messageTemplateSelfVictoryDisappearance)
                    End If
                Next teamID
                
                ' ���͗D������������P�[�X
                For teamID = 1 To 5
                    If Left(seasonStatus.Item("���ߎ��{��")(teamID), 1) = "��" And Left(seasonStatus.Item("����BA")(teamID), 1) <> "��" And Left(seasonStatus.Item("����BB")(teamID), 1) <> "��" Then
                        bodyMPBNewsOfThisSection = addLineToText(bodyMPBNewsOfThisSection, "��" & DICT_TEAM_NAME.Item(dictTeamID.Item(teamID)) & "��" & headerOfNextGame & messageTemplateSelfVictoryReappearance)
                    End If
                Next teamID
                
                ' �}�W�b�N���_������P�[�X
                For teamID = 1 To 5
                    If Left(seasonStatus.Item("���ߎ��{��")(teamID), 1) <> "M" And Left(seasonStatus.Item("����BA")(teamID), 1) = "M" And Left(seasonStatus.Item("����BB")(teamID), 1) = "M" Then
                        bodyMPBNewsOfThisSection = addLineToText(bodyMPBNewsOfThisSection, "��" & DICT_TEAM_NAME.Item(dictTeamID.Item(teamID)) & "��" & headerOfNextGame & messageTemplateMagicAppearance)
                    End If
                Next teamID
            End If
            
            ' ����_A
            headerOfNextGame = "���� " & teamOfNextSection(2, 1) & "��-��" & teamOfNextSection(2, 2) & " �ŁA"
            If seasonStatus.Item("����AA")(0) <> "" And seasonStatus.Item("����BA")(0) <> "" And seasonStatus.Item("����AA")(0) = seasonStatus.Item("����BA")(0) Then
                ' �D���`�[�������܂�P�[�X
                bodyMPBNewsOfThisSection = addLineToText(bodyMPBNewsOfThisSection, "��" & seasonStatus.Item("����BA")(0) & "��" & headerOfNextGame & messageTemplateVictory)
            Else
                ' �}�W�b�N�����ł���P�[�X
                For teamID = 1 To 5
                    If Left(seasonStatus.Item("���ߎ��{��")(teamID), 1) = "M" And Left(seasonStatus.Item("����AA")(teamID), 1) <> "M" And Left(seasonStatus.Item("����BA")(teamID), 1) <> "M" Then
                        bodyMPBNewsOfThisSection = addLineToText(bodyMPBNewsOfThisSection, "��" & DICT_TEAM_NAME.Item(dictTeamID.Item(teamID)) & "��" & headerOfNextGame & messageTemplateMagicDisappearance)
                    End If
                Next teamID
                
                ' ���͗D�������ł���P�[�X
                For teamID = 1 To 5
                    If Left(seasonStatus.Item("���ߎ��{��")(teamID), 1) <> "��" And Left(seasonStatus.Item("����AA")(teamID), 1) = "��" And Left(seasonStatus.Item("����BA")(teamID), 1) = "��" Then
                        bodyMPBNewsOfThisSection = addLineToText(bodyMPBNewsOfThisSection, "��" & DICT_TEAM_NAME.Item(dictTeamID.Item(teamID)) & "��" & headerOfNextGame & messageTemplateSelfVictoryDisappearance)
                    End If
                Next teamID
                
                ' ���͗D������������P�[�X
                For teamID = 1 To 5
                    If Left(seasonStatus.Item("���ߎ��{��")(teamID), 1) = "��" And Left(seasonStatus.Item("����AA")(teamID), 1) <> "��" And Left(seasonStatus.Item("����BA")(teamID), 1) <> "��" Then
                        bodyMPBNewsOfThisSection = addLineToText(bodyMPBNewsOfThisSection, "��" & DICT_TEAM_NAME.Item(dictTeamID.Item(teamID)) & "��" & headerOfNextGame & messageTemplateSelfVictoryReappearance)
                    End If
                Next teamID
                
                ' �}�W�b�N���_������P�[�X
                For teamID = 1 To 5
                    If Left(seasonStatus.Item("���ߎ��{��")(teamID), 1) <> "M" And Left(seasonStatus.Item("����AA")(teamID), 1) = "M" And Left(seasonStatus.Item("����BA")(teamID), 1) = "M" Then
                        bodyMPBNewsOfThisSection = addLineToText(bodyMPBNewsOfThisSection, "��" & DICT_TEAM_NAME.Item(dictTeamID.Item(teamID)) & "��" & headerOfNextGame & messageTemplateMagicAppearance)
                    End If
                Next teamID
            End If
            
            ' ����_B
            headerOfNextGame = "���� " & teamOfNextSection(2, 1) & "��-��" & teamOfNextSection(2, 2) & " �ŁA"
            If seasonStatus.Item("����AB")(0) <> "" And seasonStatus.Item("����BB")(0) <> "" And seasonStatus.Item("����AB")(0) = seasonStatus.Item("����BB")(0) Then
                ' �D���`�[�������܂�P�[�X
                bodyMPBNewsOfThisSection = addLineToText(bodyMPBNewsOfThisSection, "��" & seasonStatus.Item("����AB")(0) & "��" & headerOfNextGame & messageTemplateVictory)
            Else
                ' �}�W�b�N�����ł���P�[�X
                For teamID = 1 To 5
                    If Left(seasonStatus.Item("���ߎ��{��")(teamID), 1) = "M" And Left(seasonStatus.Item("����AB")(teamID), 1) <> "M" And Left(seasonStatus.Item("����BB")(teamID), 1) <> "M" Then
                        bodyMPBNewsOfThisSection = addLineToText(bodyMPBNewsOfThisSection, "��" & DICT_TEAM_NAME.Item(dictTeamID.Item(teamID)) & "��" & headerOfNextGame & messageTemplateMagicDisappearance)
                    End If
                Next teamID
                
                ' ���͗D�������ł���P�[�X
                For teamID = 1 To 5
                    If Left(seasonStatus.Item("���ߎ��{��")(teamID), 1) <> "��" And Left(seasonStatus.Item("����AB")(teamID), 1) = "��" And Left(seasonStatus.Item("����BB")(teamID), 1) = "��" Then
                        bodyMPBNewsOfThisSection = addLineToText(bodyMPBNewsOfThisSection, "��" & DICT_TEAM_NAME.Item(dictTeamID.Item(teamID)) & "��" & headerOfNextGame & messageTemplateSelfVictoryDisappearance)
                    End If
                Next teamID
                
                ' ���͗D������������P�[�X
                For teamID = 1 To 5
                    If Left(seasonStatus.Item("���ߎ��{��")(teamID), 1) = "��" And Left(seasonStatus.Item("����AB")(teamID), 1) <> "��" And Left(seasonStatus.Item("����BB")(teamID), 1) <> "��" Then
                        bodyMPBNewsOfThisSection = addLineToText(bodyMPBNewsOfThisSection, "��" & DICT_TEAM_NAME.Item(dictTeamID.Item(teamID)) & "��" & headerOfNextGame & messageTemplateSelfVictoryReappearance)
                    End If
                Next teamID
                
                ' �}�W�b�N���_������P�[�X
                For teamID = 1 To 5
                    If Left(seasonStatus.Item("���ߎ��{��")(teamID), 1) <> "M" And Left(seasonStatus.Item("����AB")(teamID), 1) = "M" And Left(seasonStatus.Item("����BB")(teamID), 1) = "M" Then
                        bodyMPBNewsOfThisSection = addLineToText(bodyMPBNewsOfThisSection, "��" & DICT_TEAM_NAME.Item(dictTeamID.Item(teamID)) & "��" & headerOfNextGame & messageTemplateMagicAppearance)
                    End If
                Next teamID
            End If
            
            ' ����AA
            headerOfNextGame = "���� " & teamOfNextSection(1, 1) & "��-��" & teamOfNextSection(1, 2) & " & " & teamOfNextSection(2, 1) & "��-��" & teamOfNextSection(2, 2) & " �ŁA"
            If seasonStatus.Item("����AA")(0) <> "" And seasonStatus.Item("����AB")(0) <> seasonStatus.Item("����AA")(0) And seasonStatus.Item("����BA")(0) <> seasonStatus.Item("����AA")(0) Then
                ' �D���`�[�������܂�P�[�X
                bodyMPBNewsOfThisSection = addLineToText(bodyMPBNewsOfThisSection, "��" & seasonStatus.Item("����AA")(0) & "��" & headerOfNextGame & messageTemplateVictory)
            Else
                ' �}�W�b�N�����ł���P�[�X
                For teamID = 1 To 5
                    If Left(seasonStatus.Item("���ߎ��{��")(teamID), 1) = "M" And Left(seasonStatus.Item("����AA")(teamID), 1) <> "M" And Left(seasonStatus.Item("����AB")(teamID), 1) = "M" And Left(seasonStatus.Item("����BA")(teamID), 1) = "M" Then
                        bodyMPBNewsOfThisSection = addLineToText(bodyMPBNewsOfThisSection, "��" & DICT_TEAM_NAME.Item(dictTeamID.Item(teamID)) & "��" & headerOfNextGame & messageTemplateMagicDisappearance)
                    End If
                Next teamID
                
                ' ���͗D�������ł���P�[�X
                For teamID = 1 To 5
                    If Left(seasonStatus.Item("���ߎ��{��")(teamID), 1) <> "��" And Left(seasonStatus.Item("����AA")(teamID), 1) = "��" And Left(seasonStatus.Item("����AB")(teamID), 1) <> "��" And Left(seasonStatus.Item("����BA")(teamID), 1) <> "��" Then
                        bodyMPBNewsOfThisSection = addLineToText(bodyMPBNewsOfThisSection, "��" & DICT_TEAM_NAME.Item(dictTeamID.Item(teamID)) & "��" & headerOfNextGame & messageTemplateSelfVictoryDisappearance)
                    End If
                Next teamID
                
                ' ���͗D������������P�[�X
                For teamID = 1 To 5
                    If Left(seasonStatus.Item("���ߎ��{��")(teamID), 1) = "��" And Left(seasonStatus.Item("����AA")(teamID), 1) <> "��" And Left(seasonStatus.Item("����AB")(teamID), 1) = "��" And Left(seasonStatus.Item("����BA")(teamID), 1) = "��" Then
                        bodyMPBNewsOfThisSection = addLineToText(bodyMPBNewsOfThisSection, "��" & DICT_TEAM_NAME.Item(dictTeamID.Item(teamID)) & "��" & headerOfNextGame & messageTemplateSelfVictoryReappearance)
                    End If
                Next teamID
                
                ' �}�W�b�N���_������P�[�X
                For teamID = 1 To 5
                    If Left(seasonStatus.Item("���ߎ��{��")(teamID), 1) <> "M" And Left(seasonStatus.Item("����AA")(teamID), 1) = "M" And Left(seasonStatus.Item("����AB")(teamID), 1) <> "M" And Left(seasonStatus.Item("����BA")(teamID), 1) <> "M" Then
                        bodyMPBNewsOfThisSection = addLineToText(bodyMPBNewsOfThisSection, "��" & DICT_TEAM_NAME.Item(dictTeamID.Item(teamID)) & "��" & headerOfNextGame & messageTemplateMagicAppearance)
                    End If
                Next teamID
            End If

            ' ����BA
            headerOfNextGame = "���� " & teamOfNextSection(1, 1) & "��-��" & teamOfNextSection(1, 2) & " & " & teamOfNextSection(2, 1) & "��-��" & teamOfNextSection(2, 2) & " �ŁA"
            If seasonStatus.Item("����BA")(0) <> "" And seasonStatus.Item("����BB")(0) <> seasonStatus.Item("����BA")(0) And seasonStatus.Item("����AA")(0) <> seasonStatus.Item("����BA")(0) Then
                ' �D���`�[�������܂�P�[�X
                bodyMPBNewsOfThisSection = addLineToText(bodyMPBNewsOfThisSection, "��" & seasonStatus.Item("����BA")(0) & "��" & headerOfNextGame & messageTemplateVictory)
            Else
                ' �}�W�b�N�����ł���P�[�X
                For teamID = 1 To 5
                    If Left(seasonStatus.Item("���ߎ��{��")(teamID), 1) = "M" And Left(seasonStatus.Item("����BA")(teamID), 1) <> "M" And Left(seasonStatus.Item("����BB")(teamID), 1) = "M" And Left(seasonStatus.Item("����AA")(teamID), 1) = "M" Then
                        bodyMPBNewsOfThisSection = addLineToText(bodyMPBNewsOfThisSection, "��" & DICT_TEAM_NAME.Item(dictTeamID.Item(teamID)) & "��" & headerOfNextGame & messageTemplateMagicDisappearance)
                    End If
                Next teamID
                
                ' ���͗D�������ł���P�[�X
                For teamID = 1 To 5
                    If Left(seasonStatus.Item("���ߎ��{��")(teamID), 1) <> "��" And Left(seasonStatus.Item("����BA")(teamID), 1) = "��" And Left(seasonStatus.Item("����BB")(teamID), 1) <> "��" And Left(seasonStatus.Item("����AA")(teamID), 1) <> "��" Then
                        bodyMPBNewsOfThisSection = addLineToText(bodyMPBNewsOfThisSection, "��" & DICT_TEAM_NAME.Item(dictTeamID.Item(teamID)) & "��" & headerOfNextGame & messageTemplateSelfVictoryDisappearance)
                    End If
                Next teamID
                
                ' ���͗D������������P�[�X
                For teamID = 1 To 5
                    If Left(seasonStatus.Item("���ߎ��{��")(teamID), 1) = "��" And Left(seasonStatus.Item("����BA")(teamID), 1) <> "��" And Left(seasonStatus.Item("����BB")(teamID), 1) = "��" And Left(seasonStatus.Item("����AA")(teamID), 1) = "��" Then
                        bodyMPBNewsOfThisSection = addLineToText(bodyMPBNewsOfThisSection, "��" & DICT_TEAM_NAME.Item(dictTeamID.Item(teamID)) & "��" & headerOfNextGame & messageTemplateSelfVictoryReappearance)
                    End If
                Next teamID
                
                ' �}�W�b�N���_������P�[�X
                For teamID = 1 To 5
                    If Left(seasonStatus.Item("���ߎ��{��")(teamID), 1) <> "M" And Left(seasonStatus.Item("����BA")(teamID), 1) = "M" And Left(seasonStatus.Item("����BB")(teamID), 1) <> "M" And Left(seasonStatus.Item("����AA")(teamID), 1) <> "M" Then
                        bodyMPBNewsOfThisSection = addLineToText(bodyMPBNewsOfThisSection, "��" & DICT_TEAM_NAME.Item(dictTeamID.Item(teamID)) & "��" & headerOfNextGame & messageTemplateMagicAppearance)
                    End If
                Next teamID
            End If
            
            ' ����AB
            headerOfNextGame = "���� " & teamOfNextSection(1, 1) & "��-��" & teamOfNextSection(1, 2) & " & " & teamOfNextSection(2, 1) & "��-��" & teamOfNextSection(2, 2) & " �ŁA"
            If seasonStatus.Item("����AB")(0) <> "" And seasonStatus.Item("����AA")(0) <> seasonStatus.Item("����AB")(0) And seasonStatus.Item("����BB")(0) <> seasonStatus.Item("����AB")(0) Then
                ' �D���`�[�������܂�P�[�X
                bodyMPBNewsOfThisSection = addLineToText(bodyMPBNewsOfThisSection, "��" & seasonStatus.Item("����AB")(0) & "��" & headerOfNextGame & messageTemplateVictory)
            Else
                ' �}�W�b�N�����ł���P�[�X
                For teamID = 1 To 5
                    If Left(seasonStatus.Item("���ߎ��{��")(teamID), 1) = "M" And Left(seasonStatus.Item("����AB")(teamID), 1) <> "M" And Left(seasonStatus.Item("����AA")(teamID), 1) = "M" And Left(seasonStatus.Item("����BB")(teamID), 1) = "M" Then
                        bodyMPBNewsOfThisSection = addLineToText(bodyMPBNewsOfThisSection, "��" & DICT_TEAM_NAME.Item(dictTeamID.Item(teamID)) & "��" & headerOfNextGame & messageTemplateMagicDisappearance)
                    End If
                Next teamID
                
                ' ���͗D�������ł���P�[�X
                For teamID = 1 To 5
                    If Left(seasonStatus.Item("���ߎ��{��")(teamID), 1) <> "��" And Left(seasonStatus.Item("����AB")(teamID), 1) = "��" And Left(seasonStatus.Item("����AA")(teamID), 1) <> "��" And Left(seasonStatus.Item("����BB")(teamID), 1) <> "��" Then
                        bodyMPBNewsOfThisSection = addLineToText(bodyMPBNewsOfThisSection, "��" & DICT_TEAM_NAME.Item(dictTeamID.Item(teamID)) & "��" & headerOfNextGame & messageTemplateSelfVictoryDisappearance)
                    End If
                Next teamID
                
                ' ���͗D������������P�[�X
                For teamID = 1 To 5
                    If Left(seasonStatus.Item("���ߎ��{��")(teamID), 1) = "��" And Left(seasonStatus.Item("����AB")(teamID), 1) <> "��" And Left(seasonStatus.Item("����AA")(teamID), 1) = "��" And Left(seasonStatus.Item("����BB")(teamID), 1) = "��" Then
                        bodyMPBNewsOfThisSection = addLineToText(bodyMPBNewsOfThisSection, "��" & DICT_TEAM_NAME.Item(dictTeamID.Item(teamID)) & "��" & headerOfNextGame & messageTemplateSelfVictoryReappearance)
                    End If
                Next teamID
                
                ' �}�W�b�N���_������P�[�X
                For teamID = 1 To 5
                    If Left(seasonStatus.Item("���ߎ��{��")(teamID), 1) <> "M" And Left(seasonStatus.Item("����AB")(teamID), 1) = "M" And Left(seasonStatus.Item("����AA")(teamID), 1) <> "M" And Left(seasonStatus.Item("����BB")(teamID), 1) <> "M" Then
                        bodyMPBNewsOfThisSection = addLineToText(bodyMPBNewsOfThisSection, "��" & DICT_TEAM_NAME.Item(dictTeamID.Item(teamID)) & "��" & headerOfNextGame & messageTemplateMagicAppearance)
                    End If
                Next teamID
            End If
            
            ' ����BB
            headerOfNextGame = "���� " & teamOfNextSection(1, 1) & "��-��" & teamOfNextSection(1, 2) & " & " & teamOfNextSection(2, 1) & "��-��" & teamOfNextSection(2, 2) & " �ŁA"
            If seasonStatus.Item("����BB")(0) <> "" And seasonStatus.Item("����BA")(0) <> seasonStatus.Item("����BB")(0) And seasonStatus.Item("����AB")(0) <> seasonStatus.Item("����BB")(0) Then
                ' �D���`�[�������܂�P�[�X
                bodyMPBNewsOfThisSection = addLineToText(bodyMPBNewsOfThisSection, "��" & seasonStatus.Item("����BB")(0) & "��" & headerOfNextGame & messageTemplateVictory)
            Else
                ' �}�W�b�N�����ł���P�[�X
                For teamID = 1 To 5
                    If Left(seasonStatus.Item("���ߎ��{��")(teamID), 1) = "M" And Left(seasonStatus.Item("����BB")(teamID), 1) <> "M" And Left(seasonStatus.Item("����BA")(teamID), 1) = "M" And Left(seasonStatus.Item("����AB")(teamID), 1) = "M" Then
                        bodyMPBNewsOfThisSection = addLineToText(bodyMPBNewsOfThisSection, "��" & DICT_TEAM_NAME.Item(dictTeamID.Item(teamID)) & "��" & headerOfNextGame & messageTemplateMagicDisappearance)
                    End If
                Next teamID
                
                ' ���͗D�������ł���P�[�X
                For teamID = 1 To 5
                    If Left(seasonStatus.Item("���ߎ��{��")(teamID), 1) <> "��" And Left(seasonStatus.Item("����BB")(teamID), 1) = "��" And Left(seasonStatus.Item("����BA")(teamID), 1) <> "��" And Left(seasonStatus.Item("����AB")(teamID), 1) <> "��" Then
                        bodyMPBNewsOfThisSection = addLineToText(bodyMPBNewsOfThisSection, "��" & DICT_TEAM_NAME.Item(dictTeamID.Item(teamID)) & "��" & headerOfNextGame & messageTemplateSelfVictoryDisappearance)
                    End If
                Next teamID
                
                ' ���͗D������������P�[�X
                For teamID = 1 To 5
                    If Left(seasonStatus.Item("���ߎ��{��")(teamID), 1) = "��" And Left(seasonStatus.Item("����BB")(teamID), 1) <> "��" And Left(seasonStatus.Item("����BA")(teamID), 1) = "��" And Left(seasonStatus.Item("����AB")(teamID), 1) = "��" Then
                        bodyMPBNewsOfThisSection = addLineToText(bodyMPBNewsOfThisSection, "��" & DICT_TEAM_NAME.Item(dictTeamID.Item(teamID)) & "��" & headerOfNextGame & messageTemplateSelfVictoryReappearance)
                    End If
                Next teamID
                
                ' �}�W�b�N���_������P�[�X
                For teamID = 1 To 5
                    If Left(seasonStatus.Item("���ߎ��{��")(teamID), 1) <> "M" And Left(seasonStatus.Item("����BB")(teamID), 1) = "M" And Left(seasonStatus.Item("����BA")(teamID), 1) <> "M" And Left(seasonStatus.Item("����AB")(teamID), 1) <> "M" Then
                        bodyMPBNewsOfThisSection = addLineToText(bodyMPBNewsOfThisSection, "��" & DICT_TEAM_NAME.Item(dictTeamID.Item(teamID)) & "��" & headerOfNextGame & messageTemplateMagicAppearance)
                    End If
                Next teamID
            End If
        End If
    End If

    ' ���ʂ̏o��
    If Not isDebugMode Then
        Call saveTxtFile(bodyMPBNewsOfThisSection, MPB_WORK_DIRECTORY_PATH & "\batch-min\mpbnews-section.txt")
    Else
        Call showMessageInfo(bodyMPBNewsOfThisSection, "makeMPBNewsOfThisSection")
        Call saveTxtFile(bodyMPBNewsOfThisSection, LOCAL_WORK_DIRECTORY_PATH & "\batch-min\mpbnews-section.txt")
    End If

End Function

Function seasonStatusOfSection(sectionNumber As Integer, score1D As String, score1F As String, score1H As String, score2D As String, score2F As String, score2H As String) As String()

    Dim tmp(2, 3) As String
    Dim resultArray(5) As String
        
    tmp(1, 1) = Sheets(season & "_�X�P�W���[��").Cells(sectionNumber * 8 - 5, "D").Value
    tmp(1, 2) = Sheets(season & "_�X�P�W���[��").Cells(sectionNumber * 8 - 5, "F").Value
    tmp(1, 3) = Sheets(season & "_�X�P�W���[��").Cells(sectionNumber * 8 - 5, "H").Value
    tmp(2, 1) = Sheets(season & "_�X�P�W���[��").Cells(sectionNumber * 8 - 1, "D").Value
    tmp(2, 2) = Sheets(season & "_�X�P�W���[��").Cells(sectionNumber * 8 - 1, "F").Value
    tmp(2, 3) = Sheets(season & "_�X�P�W���[��").Cells(sectionNumber * 8 - 1, "H").Value
    
    Sheets(season & "_�X�P�W���[��").Cells(sectionNumber * 8 - 5, "D").Value = score1D
    Sheets(season & "_�X�P�W���[��").Cells(sectionNumber * 8 - 5, "F").Value = score1F
    Sheets(season & "_�X�P�W���[��").Cells(sectionNumber * 8 - 5, "H").Value = score1H
    Sheets(season & "_�X�P�W���[��").Cells(sectionNumber * 8 - 1, "D").Value = score2D
    Sheets(season & "_�X�P�W���[��").Cells(sectionNumber * 8 - 1, "F").Value = score2F
    Sheets(season & "_�X�P�W���[��").Cells(sectionNumber * 8 - 1, "H").Value = score2H
    
    Application.Calculate

    Dim teamID As Integer
    resultArray(0) = ""
    For teamID = 1 To 5

        resultArray(teamID) = "-"

        If Sheets(season & "_�e��L�^").Cells(teamID + 1, "BR").Value = 0 Then
            resultArray(teamID) = "����V����"
        ElseIf Sheets(season & "_�e��L�^").Cells(teamID + 1, "BX").Value = "�D��" Then
            resultArray(teamID) = Sheets(season & "_�e��L�^").Cells(teamID + 1, "BX").Value
            resultArray(0) = DICT_TEAM_NAME.Item(dictTeamID.Item(teamID))
        ElseIf Sheets(season & "_�e��L�^").Cells(teamID + 1, "BX").Value <> "-" Then
            resultArray(teamID) = Sheets(season & "_�e��L�^").Cells(teamID + 1, "BX").Value
        End If

    Next teamID
    
    Sheets(season & "_�X�P�W���[��").Cells(sectionNumber * 8 - 5, "D").Value = tmp(1, 1)
    Sheets(season & "_�X�P�W���[��").Cells(sectionNumber * 8 - 5, "F").Value = tmp(1, 2)
    Sheets(season & "_�X�P�W���[��").Cells(sectionNumber * 8 - 5, "H").Value = tmp(1, 3)
    Sheets(season & "_�X�P�W���[��").Cells(sectionNumber * 8 - 1, "D").Value = tmp(2, 1)
    Sheets(season & "_�X�P�W���[��").Cells(sectionNumber * 8 - 1, "F").Value = tmp(2, 2)
    Sheets(season & "_�X�P�W���[��").Cells(sectionNumber * 8 - 1, "H").Value = tmp(2, 3)
    
    Application.Calculate

    seasonStatusOfSection = resultArray()

End Function

' �X�y����E���ʂ��o��
Function makeMPBNewsOfAccident()

    ' ���s����
    If section = 30 Then
        Exit Function
    End If

    ' �錾
    Dim existMPBNewsOfAccident As Boolean
    Dim bodyMPBNewsOfAccident As String
    Dim gamesBeforeThisSection As Integer
    Dim gamesAfterThisSection As Integer
    Dim teamBasedAccidentRate As Single

    ' ������
    existMPBNewsOfAccident = False
    bodyMPBNewsOfAccident = "�y�I�藣�E���z"
    bodyMPBNewsOfAccident = addLineToText(bodyMPBNewsOfAccident, "")
    gamesBeforeThisSection = -1
    gamesAfterThisSection = 0
    teamBasedAccidentRate = 0

    Dim teamID As Integer
    Dim rowIdx As Integer
    Dim columnIdx As Integer
    Dim dice As Single
    Dim visibleAccidentPeriod As Integer
    Dim hiddenAccidentPeriod As Integer
    Dim accidentInformation As String
    For teamID = 1 To 5

        ' �����󋵃`�F�b�N
        If section > 0 Then
            gamesBeforeThisSection = Sheets(season & "_�X�P�W���[��").Cells(2 + section - 1, 83 + teamID)
        End If
        gamesAfterThisSection = Sheets(season & "_�X�P�W���[��").Cells(2 + section, 83 + teamID)

        ' ��b�X�y��=(BASE_ACCIDENT_RATE)*(����a�@�K�p��)*(�����i�s�W��88.5-111.5%) ���������Ă��Ȃ��ꍇ�̓[��
        teamBasedAccidentRate = BASE_ACCIDENT_RATE * DICT_ACCIDENT_HDCP.Item(dictTeamID.Item(teamID)) * (0.885 + (gamesAfterThisSection * 0.01))
        If gamesBeforeThisSection = gamesAfterThisSection Then
            teamBasedAccidentRate = 0
        End If
        ' Call showMessageInfo(dictTeamID.Item(teamID) & " : teamBasedAccidentRate = " & teamBasedAccidentRate * 100 & "%", "makeMPBNewsOfAccident")

        ' ����X�y����
        For rowIdx = 4 + 50 * (teamID - 1) To 50 * teamID

            If Sheets(season & "_����f�[�^").Cells(rowIdx, "A").Value = "" Then
                Exit For
            End If

            ' ��b�X�y��*�X�y����W���ł̒��I �����ɃP�K���Ă���ꍇ�͑ΏۊO
            If Sheets(season & "_����f�[�^").Cells(rowIdx, 282 + gamesAfterThisSection).Value = "" Then
                Randomize
                dice = Rnd()
            Else
                dice = 1#
            End If
            ' Call showMessageInfo(dictTeamID.Item(teamID) & Sheets(season & "_����f�[�^").Cells(rowIdx, "D").Value & " : accidentRate = " & teamBasedAccidentRate * DICT_ACCIDENT_COEFFICIENT.Item(Sheets(season & "_����f�[�^").Cells(rowIdx, "E").Value) * 100 & "%, dice = " & dice * 100, "makeMPBNewsOfAccident")
            If dice < teamBasedAccidentRate * DICT_ACCIDENT_COEFFICIENT.Item(Sheets(season & "_����f�[�^").Cells(rowIdx, "E").Value) Then

                ' �X�y����(�\)���I
                visibleAccidentPeriod = drawFromDictionary(DICT_ACCIDENT_LENGTH_RATE)

                ' �X�y����(��)���I �������[���ɂ͂Ȃ�Ȃ��A������]�̏ꍇ�͕ϓ��Ȃ�
                hiddenAccidentPeriod = visibleAccidentPeriod + drawFromDictionary(DICT_ACCIDENT_MARGIN_DICT.Item(Sheets(season & "_����f�[�^").Cells(rowIdx, "E").Value))
                If hiddenAccidentPeriod = 0 Then
                    hiddenAccidentPeriod = 1
                End If
                If visibleAccidentPeriod = 24 Then
                    hiddenAccidentPeriod = 24
                End If

                ' �X�y���e���I
                accidentInformation = drawFromDictionary(DICT_ACCIDENT_INFORMATION_PITCHER_DICT.Item(visibleAccidentPeriod))
                bodyMPBNewsOfAccident = addLineToText(bodyMPBNewsOfAccident, "��" & DICT_TEAM_NAME.Item(dictTeamID.Item(teamID)) & "��" & Sheets(season & "_����f�[�^").Cells(rowIdx, "D").Value & accidentInformation)
                existMPBNewsOfAccident = True

                ' �t�@�C����������
                For columnIdx = 282 + gamesAfterThisSection To 282 + gamesAfterThisSection + hiddenAccidentPeriod - 1
                    If columnIdx > 305 Then
                        Exit For
                    End If
                    If visibleAccidentPeriod <> 24 Then
                        ' Call showMessageDebug(Sheets(season & "_����f�[�^").Cells(rowIdx, "D").Value & "(" & columnIdx - (282 + gamesAfterThisSection) + 1 & "/" & visibleAccidentPeriod & ")", "INPUT ����f�[�^.Cells(" & rowIdx & "," & columnIdx & ")")
                        Sheets(season & "_����f�[�^").Cells(rowIdx, columnIdx).Value = Sheets(season & "_����f�[�^").Cells(rowIdx, "D").Value & "(" & columnIdx - (282 + gamesAfterThisSection) + 1 & "/" & visibleAccidentPeriod & ")"
                    Else
                        ' Call showMessageDebug(Sheets(season & "_����f�[�^").Cells(rowIdx, "D").Value & "(-)", "INPUT ����f�[�^.Cells(" & rowIdx & "," & columnIdx & ")")
                        Sheets(season & "_����f�[�^").Cells(rowIdx, columnIdx).Value = Sheets(season & "_����f�[�^").Cells(rowIdx, "D").Value & "(-)"
                    End If
                Next columnIdx

            ElseIf Sheets(season & "_����f�[�^").Cells(rowIdx, 282 + gamesBeforeThisSection).Value <> "" And Sheets(season & "_����f�[�^").Cells(rowIdx, 282 + gamesAfterThisSection).Value = "" Then

                ' ���A
                bodyMPBNewsOfAccident = addLineToText(bodyMPBNewsOfAccident, "��" & DICT_TEAM_NAME.Item(dictTeamID.Item(teamID)) & "�����E����" & Sheets(season & "_����f�[�^").Cells(rowIdx, "D").Value & "�I��ɂ��āA���߂���̐�񕜋A����������܂����B")
                existMPBNewsOfAccident = True

            End If

        Next rowIdx

        ' ���X�y����
        Randomize
        For rowIdx = 4 + 50 * (teamID - 1) To 50 * teamID

            If Sheets(season & "_���f�[�^").Cells(rowIdx, "A").Value = "" Then
                Exit For
            End If

            ' ��b�X�y��*�X�y����W���ł̒��I �����ɃP�K���Ă���ꍇ�͑ΏۊO
            If Sheets(season & "_���f�[�^").Cells(rowIdx, 236 + gamesAfterThisSection).Value = "" Then
                Randomize
                dice = Rnd()
            Else
                dice = 1
            End If
            ' Call showMessageInfo(dictTeamID.Item(teamID) & Sheets(season & "_���f�[�^").Cells(rowIdx, "D").Value & " : accidentRate = " & teamBasedAccidentRate * DICT_ACCIDENT_COEFFICIENT.Item(Sheets(season & "_���f�[�^").Cells(rowIdx, "E").Value) * 100 & "%, dice = " & dice * 100, "makeMPBNewsOfAccident")
            If dice < teamBasedAccidentRate * DICT_ACCIDENT_COEFFICIENT.Item(Sheets(season & "_���f�[�^").Cells(rowIdx, "E").Value) Then

                ' �X�y����(�\)���I
                visibleAccidentPeriod = drawFromDictionary(DICT_ACCIDENT_LENGTH_RATE)

                ' �X�y����(��)���I �������[���ɂ͂Ȃ�Ȃ��A������]�̏ꍇ�͕ϓ��Ȃ�
                hiddenAccidentPeriod = visibleAccidentPeriod + drawFromDictionary(DICT_ACCIDENT_MARGIN_DICT.Item(Sheets(season & "_���f�[�^").Cells(rowIdx, "E").Value))
                If hiddenAccidentPeriod = 0 Then
                    hiddenAccidentPeriod = 1
                End If
                If visibleAccidentPeriod = 24 Then
                    hiddenAccidentPeriod = 24
                End If

                ' �X�y���e���I
                accidentInformation = drawFromDictionary(DICT_ACCIDENT_INFORMATION_FIELDER_DICT.Item(visibleAccidentPeriod))
                bodyMPBNewsOfAccident = addLineToText(bodyMPBNewsOfAccident, "��" & DICT_TEAM_NAME.Item(dictTeamID.Item(teamID)) & "��" & Sheets(season & "_���f�[�^").Cells(rowIdx, "D").Value & accidentInformation)
                existMPBNewsOfAccident = True

                ' �t�@�C����������
                For columnIdx = 236 + gamesAfterThisSection To 236 + gamesAfterThisSection + hiddenAccidentPeriod - 1
                    If columnIdx > 259 Then
                        Exit For
                    End If
                    If visibleAccidentPeriod <> 24 Then
                        ' Call showMessageDebug(Sheets(season & "_���f�[�^").Cells(rowIdx, "D").Value & "(" & columnIdx - (236 + gamesAfterThisSection) + 1 & "/" & visibleAccidentPeriod & ")", "INPUT ���f�[�^.Cells(" & rowIdx & "," & columnIdx & ")")
                        Sheets(season & "_���f�[�^").Cells(rowIdx, columnIdx).Value = Sheets(season & "_���f�[�^").Cells(rowIdx, "D").Value & "(" & columnIdx - (236 + gamesAfterThisSection) + 1 & "/" & visibleAccidentPeriod & ")"
                    Else
                        ' Call showMessageDebug(Sheets(season & "_���f�[�^").Cells(rowIdx, "D").Value & "(-)", "INPUT ���f�[�^.Cells(" & rowIdx & "," & columnIdx & ")")
                        Sheets(season & "_���f�[�^").Cells(rowIdx, columnIdx).Value = Sheets(season & "_���f�[�^").Cells(rowIdx, "D").Value & "(-)"
                    End If
                Next columnIdx

            ElseIf Sheets(season & "_���f�[�^").Cells(rowIdx, 236 + gamesBeforeThisSection).Value <> "" And Sheets(season & "_���f�[�^").Cells(rowIdx, 236 + gamesAfterThisSection).Value = "" Then

                ' ���A
                bodyMPBNewsOfAccident = addLineToText(bodyMPBNewsOfAccident, "��" & DICT_TEAM_NAME.Item(dictTeamID.Item(teamID)) & "�����E����" & Sheets(season & "_���f�[�^").Cells(rowIdx, "D").Value & "�I��ɂ��āA���߂���̐�񕜋A����������܂����B")
                existMPBNewsOfAccident = True

            End If

        Next rowIdx

    Next teamID

    ' ���ʂ̏o��
    If Not existMPBNewsOfAccident Then
        bodyMPBNewsOfAccident = addLineToText(bodyMPBNewsOfAccident, "�I��̗��E/���A�Ɋւ�����͂���܂���B")
    End If

    If Not isDebugMode Then
        Call saveTxtFile(bodyMPBNewsOfAccident, MPB_WORK_DIRECTORY_PATH & "\batch-min\mpbnews-accident.txt")
    Else
        Call showMessageInfo(bodyMPBNewsOfAccident, "makeMPBNewsOfAccident")
        Call saveTxtFile(bodyMPBNewsOfAccident, LOCAL_WORK_DIRECTORY_PATH & "\batch-min\mpbnews-accident.txt")
    End If

End Function

' ���ߓ��������̈˗����o��
Function makeMPBNewsOfNextGame()

    ' ���s����
    If section = 30 Then
        Exit Function
    End If

    ' �錾
    Dim bodyMPBNewsOfNextGame As String

    ' ������
    bodyMPBNewsOfNextGame = addLineToText(bodyMPBNewsOfNextGame, "���������̒����ɂ����͂����肢�������܂��B")
    bodyMPBNewsOfNextGame = addLineToText(bodyMPBNewsOfNextGame, "")

    bodyMPBNewsOfNextGame = addLineToText(bodyMPBNewsOfNextGame, "[��" & section + 1 & "��]")
    If Sheets(season & "_�X�P�W���[��").Cells(8 * section + 3, "F").Value <> "" Then
        bodyMPBNewsOfNextGame = addLineToText(bodyMPBNewsOfNextGame, "<���{��>�@" & Sheets(season & "_�X�P�W���[��").Cells(8 * section + 2, "C").Value & " " & Sheets(season & "_�X�P�W���[��").Cells(8 * section + 3, "D").Value & " - " & Sheets(season & "_�X�P�W���[��").Cells(8 * section + 3, "H").Value & " " & Sheets(season & "_�X�P�W���[��").Cells(8 * section + 2, "J").Value)
    Else
        bodyMPBNewsOfNextGame = addLineToText(bodyMPBNewsOfNextGame, Sheets(season & "_�X�P�W���[��").Cells(8 * section + 2, "C").Value & "(" & Sheets(season & "_�X�P�W���[��").Cells(8 * section + 2, "D").Value & ") - (" & Sheets(season & "_�X�P�W���[��").Cells(8 * section + 2, "H").Value & ") " & Sheets(season & "_�X�P�W���[��").Cells(8 * section + 2, "J").Value)
    End If
    If Sheets(season & "_�X�P�W���[��").Cells(8 * section + 7, "F").Value Then
        bodyMPBNewsOfNextGame = addLineToText(bodyMPBNewsOfNextGame, "<���{��>�@" & Sheets(season & "_�X�P�W���[��").Cells(8 * section + 6, "C").Value & " " & Sheets(season & "_�X�P�W���[��").Cells(8 * section + 7, "D").Value & " - " & Sheets(season & "_�X�P�W���[��").Cells(8 * section + 7, "H").Value & " " & Sheets(season & "_�X�P�W���[��").Cells(8 * section + 6, "J").Value)
    Else
        bodyMPBNewsOfNextGame = addLineToText(bodyMPBNewsOfNextGame, Sheets(season & "_�X�P�W���[��").Cells(8 * section + 6, "C").Value & "(" & Sheets(season & "_�X�P�W���[��").Cells(8 * section + 6, "D").Value & ") - (" & Sheets(season & "_�X�P�W���[��").Cells(8 * section + 6, "H").Value & ") " & Sheets(season & "_�X�P�W���[��").Cells(8 * section + 6, "J").Value)
    End If

    bodyMPBNewsOfNextGame = addLineToText(bodyMPBNewsOfNextGame, "")

    If section <= 28 Then
        bodyMPBNewsOfNextGame = addLineToText(bodyMPBNewsOfNextGame, "[��" & section + 2 & "��]")
        bodyMPBNewsOfNextGame = addLineToText(bodyMPBNewsOfNextGame, Sheets(season & "_�X�P�W���[��").Cells(8 * section + 10, "C").Value & "(" & Sheets(season & "_�X�P�W���[��").Cells(8 * section + 10, "D").Value & ") - (" & Sheets(season & "_�X�P�W���[��").Cells(8 * section + 10, "H").Value & ") " & Sheets(season & "_�X�P�W���[��").Cells(8 * section + 10, "J").Value)
        bodyMPBNewsOfNextGame = addLineToText(bodyMPBNewsOfNextGame, Sheets(season & "_�X�P�W���[��").Cells(8 * section + 14, "C").Value & "(" & Sheets(season & "_�X�P�W���[��").Cells(8 * section + 14, "D").Value & ") - (" & Sheets(season & "_�X�P�W���[��").Cells(8 * section + 14, "H").Value & ") " & Sheets(season & "_�X�P�W���[��").Cells(8 * section + 14, "J").Value)
    End If

    If Not isDebugMode Then
        Call saveTxtFile(bodyMPBNewsOfNextGame, MPB_WORK_DIRECTORY_PATH & "\batch-week\mpbnews-nextgame.txt")
    Else
        Call showMessageInfo(bodyMPBNewsOfNextGame, "makeMPBNewsOfNextGame")
        Call saveTxtFile(bodyMPBNewsOfNextGame, LOCAL_WORK_DIRECTORY_PATH & "\batch-week\mpbnews-nextgame.txt")
    End If

End Function

' �X�P�W���[���摜���o��
Function savePictureOfSchedule()

    Application.Calculate

    If Not isDebugMode Then
        Call savePngFile(Sheets(season & "_�X�P�W���[��").Range("A" & WorksheetFunction.Max(1, section * 8 - 6) & ":AG" & WorksheetFunction.Max(41, section * 8 - 6 + 55)), MPB_WORK_DIRECTORY_PATH & "\batch-min\mpbpicture-schedule.png")
    Else
        Call savePngFile(Sheets(season & "_�X�P�W���[��").Range("A" & WorksheetFunction.Max(1, section * 8 - 6) & ":AG" & WorksheetFunction.Max(41, section * 8 - 6 + 55)), LOCAL_WORK_DIRECTORY_PATH & "\batch-min\mpbpicture-schedule.png")
    End If


End Function

' ���щ摜���o��
Function savePictureOfRecord()

    Application.Calculate

    If Not isDebugMode Then
        Call savePngFile(Sheets(season & "_�e��L�^").Range("A1:AR41"), MPB_WORK_DIRECTORY_PATH & "\batch-min\mpbpicture-record.png")
    Else
        Call savePngFile(Sheets(season & "_�e��L�^").Range("A1:AR41"), LOCAL_WORK_DIRECTORY_PATH & "\batch-min\mpbpicture-record.png")
    End If

End Function
