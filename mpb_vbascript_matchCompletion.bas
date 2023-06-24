Attribute VB_Name = "mpb_vbascript_matchCompletion"
Option Explicit

Dim season As String
Dim game As Integer
Dim section As Integer

Dim dictTeamID As New Dictionary

Sub matchCompletion()
    
    ' �f�o�b�O���[�h
    Call DebugMode

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
    
    Call Backup
    
    season = ActiveSheet.Cells(1, "A").Value
    game = WorksheetFunction.CountIf(ActiveSheet.Range("BA2:BA241"), 0) / 4
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
Function ExitProcess()
    
    Sheets(season & "_�X�P�W���[��").Protect AllowFormattingColumns:=True, AllowFormattingRows:=True
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
    If Sheets(season & "_�X�P�W���[��").Cells(section * 8 + 3, "D").Value <> "" Or Sheets(season & "_�X�P�W���[��").Cells(section * 8 + 7, "D").Value <> "" Or _
       Sheets(season & "_�X�P�W���[��").Cells(section * 8 + 3, "F").Value <> "" Or Sheets(season & "_�X�P�W���[��").Cells(section * 8 + 7, "F").Value <> "" Or _
       Sheets(season & "_�X�P�W���[��").Cells(section * 8 + 3, "H").Value <> "" Or Sheets(season & "_�X�P�W���[��").Cells(section * 8 + 7, "H").Value <> "" Then
        Call MessageError("�s�����̓G���[", "IsSectionCompleted")
        Call ExitProcess
    End If
    
    ' �J���O�܂��͍ŏI�ߌ�ŗ\���攭���l����K�v���Ȃ��p�^�[��
    If section = 0 Or section = 30 Then
        IsSectionCompleted = True
        Exit Function
    End If
    
    ' �\���攭���o�����Ă��Ȃ��p�^�[��
    If Sheets(season & "_�X�P�W���[��").Cells(section * 8 + 2, "D").Value = "" Or Sheets(season & "_�X�P�W���[��").Cells(section * 8 + 6, "D").Value = "" Or _
       Sheets(season & "_�X�P�W���[��").Cells(section * 8 + 2, "H").Value = "" Or Sheets(season & "_�X�P�W���[��").Cells(section * 8 + 6, "H").Value = "" Then
        Call MessageError("�\���攭�������G���[", "IsSectionCompleted")
        Call ExitProcess
    End If
    
    IsSectionCompleted = True
    
End Function

' �߂̐i�s�ɂ�蔭������A���炩���ߗ\�肳��Ă���C�x���g���o��
Function MakeMPBNewsSeasonEvent()
    
    ' �錾
    Dim mpbNewsSeasonEventFlg As Boolean
    Dim mpbNewsSeasonEvent As String
    Dim tsobBorderDict As New Dictionary
    
    ' ������
    mpbNewsSeasonEventFlg = False
    mpbNewsSeasonEvent = "�yMPB�^�c����̂��m�点�z"
    
    ' TSOB�g�̐U�蒼��
    If section = 10 Or section = 20 Then
        mpbNewsSeasonEventFlg = True
        
        mpbNewsSeasonEvent = AddRowText(mpbNewsSeasonEvent, "�ETSOB�g�̐U�蒼�����s���܂��B")
        mpbNewsSeasonEvent = AddRowText(mpbNewsSeasonEvent, "- - - - - - - - - -")
        mpbNewsSeasonEvent = AddRowText(mpbNewsSeasonEvent, "1��: " & Left(Sheets(season & "_�e��L�^").Cells(2, "B").Value, 1) & " �� 3.5")
        mpbNewsSeasonEvent = AddRowText(mpbNewsSeasonEvent, "2��: " & Left(Sheets(season & "_�e��L�^").Cells(3, "B").Value, 1) & " �� 4.0")
        mpbNewsSeasonEvent = AddRowText(mpbNewsSeasonEvent, "3��: " & Left(Sheets(season & "_�e��L�^").Cells(4, "B").Value, 1) & " �� 4.5")
        mpbNewsSeasonEvent = AddRowText(mpbNewsSeasonEvent, "4��: " & Left(Sheets(season & "_�e��L�^").Cells(5, "B").Value, 1) & " �� 5.0")
        mpbNewsSeasonEvent = AddRowText(mpbNewsSeasonEvent, "5��: " & Left(Sheets(season & "_�e��L�^").Cells(6, "B").Value, 1) & " �� 5.5")
        mpbNewsSeasonEvent = AddRowText(mpbNewsSeasonEvent, "�������`�[���������ɂ́A�K���������̒ʂ�ƂȂ�Ȃ��ꍇ������܂��B")
        mpbNewsSeasonEvent = AddRowText(mpbNewsSeasonEvent, "")
        
        tsobBorderDict.Add Left(Sheets(season & "_�e��L�^").Cells(2, "B").Value, 1), "3.5"
        tsobBorderDict.Add Left(Sheets(season & "_�e��L�^").Cells(3, "B").Value, 1), "4.0"
        tsobBorderDict.Add Left(Sheets(season & "_�e��L�^").Cells(4, "B").Value, 1), "4.5"
        tsobBorderDict.Add Left(Sheets(season & "_�e��L�^").Cells(5, "B").Value, 1), "5.0"
        tsobBorderDict.Add Left(Sheets(season & "_�e��L�^").Cells(6, "B").Value, 1), "5.5"
        
        Sheets(season & "_�X�P�W���[��").Cells(27, "CP").Value = tsobBorderDict.Item(Sheets(season & "_�X�P�W���[��").Cells(1, "BB").Value)
        Sheets(season & "_�X�P�W���[��").Cells(27, "CQ").Value = tsobBorderDict.Item(Sheets(season & "_�X�P�W���[��").Cells(1, "BC").Value)
        Sheets(season & "_�X�P�W���[��").Cells(27, "CR").Value = tsobBorderDict.Item(Sheets(season & "_�X�P�W���[��").Cells(1, "BD").Value)
        Sheets(season & "_�X�P�W���[��").Cells(27, "CS").Value = tsobBorderDict.Item(Sheets(season & "_�X�P�W���[��").Cells(1, "BE").Value)
        Sheets(season & "_�X�P�W���[��").Cells(27, "CT").Value = tsobBorderDict.Item(Sheets(season & "_�X�P�W���[��").Cells(1, "BF").Value)
    End If
    
    ' HDCP�ύX��t�J�n
    If section = 10 Or section = 20 Then
        mpbNewsSeasonEventFlg = True
        
        mpbNewsSeasonEvent = AddRowText(mpbNewsSeasonEvent, "�E�������A�㔼�킩���HDCP�ύX��t���J�n���܂��B��15�ߏI���������Ē��ߐ؂�̂ŁA�ύX�������`�[���́A�K�v�ɉ����Đ\�����s���Ă��������B�ύX���Ȃ��ꍇ�́A���ɑΉ��s�v�ł��B")
        mpbNewsSeasonEvent = AddRowText(mpbNewsSeasonEvent, "")
    End If
    
    ' HDCP�ύX��
    If section = 11 Or section = 12 Or section = 13 Or section = 14 Then
        mpbNewsSeasonEventFlg = True
        
        mpbNewsSeasonEvent = AddRowText(mpbNewsSeasonEvent, "�E�㔼�킩���HDCP�ύX����t���ł��B�ύX�������`�[���́A��15�ߏI���܂łɐ\�����s���Ă��������B")
        mpbNewsSeasonEvent = AddRowText(mpbNewsSeasonEvent, "")
    End If
    
    ' HDCP�ύX��t�Y
    If section = 15 Then
        mpbNewsSeasonEventFlg = True
        
        mpbNewsSeasonEvent = AddRowText(mpbNewsSeasonEvent, "�E�����������܂��āA�㔼��Ɍ�����HDCP�ύX�̐\������ߐ؂�܂��BHDCP�̕\���ݒ���ŐV�����Ă��������B")
        mpbNewsSeasonEvent = AddRowText(mpbNewsSeasonEvent, "")
    End If
    
    ' B9GG��o��t�J�n
    If section = 25 Then
        mpbNewsSeasonEventFlg = True
        
        mpbNewsSeasonEvent = AddRowText(mpbNewsSeasonEvent, "�E�������AB9GG�m�~�l�[�g�I�[�_�[�̒�o��t���J�n���܂��B��28�ߏI���������Ē��ߐ؂�̂ŁA�e�`�[���ALINE�O���[�v�̃A���o���u" & season & "B9GG�m�~�l�[�g�v�ɒ�o�����肢�������܂��B")
        mpbNewsSeasonEvent = AddRowText(mpbNewsSeasonEvent, "")
    End If
    
    ' B9GG��o��t��
    If section = 26 Or section = 27 Then
        mpbNewsSeasonEventFlg = True
        
        mpbNewsSeasonEvent = AddRowText(mpbNewsSeasonEvent, "�EB9GG�m�~�l�[�g�I�[�_�[�̒�o/�ύX����t���ł��B����o�̃`�[���́A��28�߂��I������܂łɁALINE�O���[�v�̃A���o���u" & season & "B9GG�m�~�l�[�g�v�ւ̒�o�����肢�������܂��B")
        mpbNewsSeasonEvent = AddRowText(mpbNewsSeasonEvent, "")
    End If
    
    ' B9GG��o��t�Y
    If section = 28 Then
        mpbNewsSeasonEventFlg = True
        
        mpbNewsSeasonEvent = AddRowText(mpbNewsSeasonEvent, "�E�����������܂��āAB9GG�m�~�l�[�g�I�[�_�[�̒�o����ߐ؂�܂��B")
        mpbNewsSeasonEvent = AddRowText(mpbNewsSeasonEvent, "")
    End If
    
    ' MPB�A���[�h�ē�
    If section = 30 Then
        mpbNewsSeasonEventFlg = True
        
        mpbNewsSeasonEvent = AddRowText(mpbNewsSeasonEvent, "�E���V�[�Y���A�\�肳��Ă����S�������I�����܂����B�܂��́A�F���񂨔�ꂳ�܂ł����I���̌�AMPB�A���[�h�����{���܂��̂ŁA�ē������҂����������B")
        mpbNewsSeasonEvent = AddRowText(mpbNewsSeasonEvent, "")
    End If
    
    ' ���ʂ̏o��
    If mpbNewsSeasonEventFlg Then
        mpbNewsSeasonEvent = AddRowText(mpbNewsSeasonEvent, "�ȏ�")
        
        If Not debugModeFlg Then
            Call OutputText(mpbNewsSeasonEvent, MPB_WORK_DIRECTORY_PATH & "\mpbnews-seasonevent.txt")
        Else
            Call MessageInfo(mpbNewsSeasonEvent, "MakeMPBNewsSeasonEvent")
        End If
    End If

End Function

' �߂̐i�s�ɂ�蔭������A�D���}�W�b�N�⎩�͗D���Ɋւ���C�x���g���o��
Function MakeMPBNewsOfThisSection()
    
    ' ���s����
    If section = 0 Or True Then
        Exit Function
    End If
    
    ' �錾
    Dim mpbNewsOfThisSectionFlg As Boolean
    Dim mpbNewsOfThisSection As String
    Dim seasonStatus As New Dictionary
    
    ' ������
    mpbNewsOfThisSectionFlg = False
    mpbNewsOfThisSection = "�yMPB�j���[�X�z"
    
    ' �󋵊m�F(���ߎ��{�O)
    seasonStatus.Add "���ߎ��{�O", CheckSeasonStatus(section - 1, [["","",""],["","",""]])
    
    ' ���ߎ��{�O�ɗD�������܂��Ă���ꍇ�̓X�L�b�v
    If seasonStatus.Item("���ߎ��{�O")(0) <> "" Then
        Exit Function
    End If
    
    ' �󋵊m�F(���ߎ��{��)
    seasonStatus.Add "���ߎ��{��", CheckSeasonStatus(section, [["","",""],["","",""]])
    
    ' ���߂��l����K�v���Ȃ��ꍇ
    If seasonStatus.Item("���ߎ��{��")(0) <> "" Or section = 30 Then
        Dim teamID As Integer
        For teamID = 1 To 5
            If seasonStatus.Item("���ߎ��{��")(teamID) = "�D��" Then
                mpbNewsOfThisSectionFlg = True
                mpbNewsOfThisSection = AddRowText(mpbNewsOfThisSection, "��" & DICT_TEAMNAME.Item(dictTeamID.Item(teamID)) & "��MPB(" & season & ")�D�����m��I")
            End If
        Next teamID
    End If
    
    ' �󋵊m�F(���ߎ��{��)
    If Not mpbNewsOfThisSectionFlg Then
        seasonStatus.Add "���߁�-��/��-��", CheckSeasonStatus(section + 1, [["9","-","0"],["9","-","0"]])
        seasonStatus.Add "���߁�-��/��-��", CheckSeasonStatus(section + 1, [["9","-","0"],["0","-","9"]])
        seasonStatus.Add "���߁�-��/��-��", CheckSeasonStatus(section + 1, [["0","-","9"],["9","-","0"]])
        seasonStatus.Add "���߁�-��/��-��", CheckSeasonStatus(section + 1, [["0","-","9"],["0","-","9"]])
    End If
    
    ' Coming Soon
    
    ' ���ʂ̏o��
    If mpbNewsOfThisSectionFlg Then
        If Not debugModeFlg Then
            Call OutputText(mpbNewsOfThisSection, MPB_WORK_DIRECTORY_PATH & "\mpbnews-section.txt")
        Else
            Call MessageInfo(mpbNewsOfThisSection, "MakeMPBNewsOfThisSection")
        End If
    End If
    
End Function

Function CheckSeasonStatus(sectionNumber As Integer, ByRef score As String) As String()
    
    Dim tmp(2) As String
    Dim resultArray(5) As String
    
    If sectionNumber < section Then
        tmp(1) = Sheets(season & "_�X�P�W���[��").Cells(sectionNumber * 8 + 3, "F").Value
        tmp(2) = Sheets(season & "_�X�P�W���[��").Cells(sectionNumber * 8 + 7, "F").Value
        Sheets(season & "_�X�P�W���[��").Cells(sectionNumber * 8 + 3, "F").Value = ""
        Sheets(season & "_�X�P�W���[��").Cells(sectionNumber * 8 + 7, "F").Value = ""
    ElseIf sectionNumber > section Then
        Sheets(season & "_�X�P�W���[��").Cells(sectionNumber * 8 + 3, "D").Value = score(0, 0)
        Sheets(season & "_�X�P�W���[��").Cells(sectionNumber * 8 + 3, "F").Value = score(0, 1)
        Sheets(season & "_�X�P�W���[��").Cells(sectionNumber * 8 + 3, "H").Value = score(0, 2)
        Sheets(season & "_�X�P�W���[��").Cells(sectionNumber * 8 + 7, "D").Value = score(1, 0)
        Sheets(season & "_�X�P�W���[��").Cells(sectionNumber * 8 + 7, "F").Value = score(1, 1)
        Sheets(season & "_�X�P�W���[��").Cells(sectionNumber * 8 + 7, "H").Value = score(1, 2)
    End If
    
    Application.Calculate
    
    Dim teamID As Integer
    resultArray(0) = ""
    For teamID = 1 To 5
        
        resultArray(teamID) = "-"
        
        If Sheets(seasonName & "_�e��L�^").Cells(teamID + 1, "BR").Value = 0 Then
            resultArray(teamID) = "����V����"
        ElseIf Sheets(seasonName & "_�e��L�^").Cells(teamID + 1, "BX").Value = "�D��" Then
            resultArray(teamID) = Sheets(seasonName & "_�e��L�^").Cells(teamID + 1, "BX").Value
            resultArray(0) = "�D���`�[������"
        ElseIf Sheets(seasonName & "_�e��L�^").Cells(teamID + 1, "BX").Value <> "-" Then
            resultArray(teamID) = Sheets(seasonName & "_�e��L�^").Cells(teamID + 1, "BX").Value
        End If
        
    Next teamID
    
    If sectionNumber < section Then
        Sheets(season & "_�X�P�W���[��").Cells(sectionNumber * 8 + 3, "F").Value = tmp(1)
        Sheets(season & "_�X�P�W���[��").Cells(sectionNumber * 8 + 7, "F").Value = tmp(2)
    ElseIf sectionNumber > section Then
        Sheets(season & "_�X�P�W���[��").Cells(sectionNumber * 8 + 3, "D").Value = ""
        Sheets(season & "_�X�P�W���[��").Cells(sectionNumber * 8 + 3, "F").Value = ""
        Sheets(season & "_�X�P�W���[��").Cells(sectionNumber * 8 + 3, "H").Value = ""
        Sheets(season & "_�X�P�W���[��").Cells(sectionNumber * 8 + 7, "D").Value = ""
        Sheets(season & "_�X�P�W���[��").Cells(sectionNumber * 8 + 7, "F").Value = ""
        Sheets(season & "_�X�P�W���[��").Cells(sectionNumber * 8 + 7, "H").Value = ""
    End If
    
    Application.Calculate
    
    CheckSeasonStatus = resultArray()
    
End Function

' �X�y����E���ʂ��o��
Function MakeMPBNewsOfAccident()
    
    ' ���s����
    If section = 30 Then
        Exit Function
    End If
    
    ' �錾
    Dim mpbNewsOfAccidentFlg As Boolean
    Dim mpbNewsOfAccident As String
    Dim gamesBeforeThisSection As Integer
    Dim gamesAfterThisSection As Integer
    Dim teamBasedAccidentRate As Single
    
    ' ������
    mpbNewsOfAccidentFlg = False
    mpbNewsOfAccident = "�yMPB�j���[�X�z"
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
    Dim accidentInformationFile As String
    Dim accidentInformationNews As String
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
        Call MessageInfo(dictTeamID.Item(teamID) & " : teamBasedAccidentRate = " & teamBasedAccidentRate * 100 & "%", "MakeMPBNewsOfAccident")
        
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
                dice = 1
            End If
            If dice < teamBasedAccidentRate * DICT_ACCIDENT_COEFFICIENT.Item(Sheets(season & "_����f�[�^").Cells(rowIdx, "E").Value) Then
                
                ' �X�y����(�\)���I
                visibleAccidentPeriod = DrawFromDict(DICT_ACCIDENT_LENGTH_RATE)
                
                ' �X�y����(��)���I �������[���ɂ͂Ȃ�Ȃ�
                hiddenAccidentPeriod = visibleAccidentPeriod + DrawFromDict(DICT_ACCIDENT_MARGIN_DICT.Item(Sheets(season & "_����f�[�^").Cells(rowIdx, "E").Value))
                If hiddenAccidentPeriod = 0 Then
                    hiddenAccidentPeriod = 1
                End If
                
                ' �X�y���e���I
                accidentInformation = DrawFromDict(DICT_ACCIDENT_INFORMATION_PITCHER_DICT.Item(visibleAccidentPeriod))
                accidentInformationFile = Split(accidentInformation, "_")(0)
                accidentInformationNews = Split(accidentInformation, "_")(1)
                mpbNewsOfAccident = AddRowText(mpbNewsOfAccident, "��" & DICT_TEAMNAME.Item(dictTeamID.Item(teamID)) & "��" & Sheets(season & "_����f�[�^").Cells(rowIdx, "D").Value & "�I�肪" & accidentInformationNews)
                mpbNewsOfAccidentFlg = True
                
                ' �t�@�C����������
                For columnIdx = 282 + gamesAfterThisSection To 282 + gamesAfterThisSection + hiddenAccidentPeriod - 1
                    If columnIdx > 305 Then
                        Exit For
                    End If
                    Call MessageDebug(Sheets(season & "_����f�[�^").Cells(rowIdx, "D").Value & ":" & accidentInformationFile & "(" & visibleAccidentPeriod & ")", "INPUT ����f�[�^.Cells(" & rowIdx & "," & columnIdx & ")")
                    Sheets(season & "_����f�[�^").Cells(rowIdx, columnIdx).Value = Sheets(season & "_����f�[�^").Cells(rowIdx, "D").Value & ":" & accidentInformationFile & "(" & visibleAccidentPeriod & ")"
                Next columnIdx
            
            ElseIf Sheets(season & "_����f�[�^").Cells(rowIdx, 282 + gamesBeforeThisSection).Value <> "" And Sheets(season & "_����f�[�^").Cells(rowIdx, 282 + gamesAfterThisSection).Value = "" Then
                
                ' ���A
                mpbNewsOfAccident = AddRowText(mpbNewsOfAccident, "��" & DICT_TEAMNAME.Item(dictTeamID.Item(teamID)) & "�����E����" & Sheets(season & "_����f�[�^").Cells(rowIdx, "D").Value & "�I��ɂ��āA���߂���̐�񕜋A����������܂����B")
                mpbNewsOfAccidentFlg = True
                
            End If
            
        Next rowIdx
        
        ' ���X�y����
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
            If dice < teamBasedAccidentRate * DICT_ACCIDENT_COEFFICIENT.Item(Sheets(season & "_���f�[�^").Cells(rowIdx, "E").Value) Then
                
                ' �X�y����(�\)���I
                visibleAccidentPeriod = DrawFromDict(DICT_ACCIDENT_LENGTH_RATE)
                
                ' �X�y����(��)���I �������[���ɂ͂Ȃ�Ȃ�
                hiddenAccidentPeriod = visibleAccidentPeriod + DrawFromDict(DICT_ACCIDENT_MARGIN_DICT.Item(Sheets(season & "_���f�[�^").Cells(rowIdx, "E").Value))
                If hiddenAccidentPeriod = 0 Then
                    hiddenAccidentPeriod = 1
                End If
                
                ' �X�y���e���I
                accidentInformation = DrawFromDict(DICT_ACCIDENT_INFORMATION_FIELDER_DICT.Item(visibleAccidentPeriod))
                accidentInformationFile = Split(accidentInformation, "_")(0)
                accidentInformationNews = Split(accidentInformation, "_")(1)
                mpbNewsOfAccident = AddRowText(mpbNewsOfAccident, "��" & DICT_TEAMNAME.Item(dictTeamID.Item(teamID)) & "��" & Sheets(season & "_���f�[�^").Cells(rowIdx, "D").Value & "�I�肪" & accidentInformationNews)
                mpbNewsOfAccidentFlg = True
                
                ' �t�@�C����������
                For columnIdx = 236 + gamesAfterThisSection To 236 + gamesAfterThisSection + hiddenAccidentPeriod - 1
                    If columnIdx > 259 Then
                        Exit For
                    End If
                    Call MessageDebug(Sheets(season & "_���f�[�^").Cells(rowIdx, "D").Value & ":" & accidentInformationFile & "(" & visibleAccidentPeriod & ")", "INPUT ���f�[�^.Cells(" & rowIdx & "," & columnIdx & ")")
                    Sheets(season & "_���f�[�^").Cells(rowIdx, columnIdx).Value = Sheets(season & "_���f�[�^").Cells(rowIdx, "D").Value & ":" & accidentInformationFile & "(" & visibleAccidentPeriod & ")"
                Next columnIdx
                
                
            ElseIf Sheets(season & "_���f�[�^").Cells(rowIdx, 236 + gamesBeforeThisSection).Value <> "" And Sheets(season & "_���f�[�^").Cells(rowIdx, 236 + gamesAfterThisSection).Value = "" Then

                ' ���A
                mpbNewsOfAccident = AddRowText(mpbNewsOfAccident, "��" & DICT_TEAMNAME.Item(dictTeamID.Item(teamID)) & "�����E����" & Sheets(season & "_���f�[�^").Cells(rowIdx, "D").Value & "�I��ɂ��āA���߂���̐�񕜋A����������܂����B")
                mpbNewsOfAccidentFlg = True
                
            End If
        
        Next rowIdx
        
    Next teamID
    
    ' ���ʂ̏o��
    If mpbNewsOfAccidentFlg Then
        If Not debugModeFlg Then
            Call OutputText(mpbNewsOfAccident, MPB_WORK_DIRECTORY_PATH & "\mpbnews-accident.txt")
        Else
            Call MessageInfo(mpbNewsOfAccident, "MakeMPBNewsOfAccident")
        End If
    End If
    
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

