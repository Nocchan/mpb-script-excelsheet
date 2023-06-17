Attribute VB_Name = "mpb_vbascript_sectionCompletion"
Sub �A�N�V�f���g����()

    Dim debugModeFlg As Boolean
    debugModeFlg = False
    If debugModeFlg Then
        MsgBox "�f�o�b�O���[�h"
    End If
    
    ' �G���[�`�F�b�N
    If ActiveSheet.Name <> ActiveSheet.Cells(1, "A").Value & "_�X�P�W���[��" Then
        MsgBox "�V�[�g���܂���A1�Z���̃V�[�Y���w�肪�s���ł��B"
        End
    End If
    
    If Not debugModeFlg Then
        Application.ScreenUpdating = False
    End If
    
    ' �S�̂Ŏg�p����ϐ�
    Dim seasonName As String
    Dim numberOfSection As Integer
    Dim allAccidentResultNotification, tsobChangeNotification, rankNotification As String
    
    ' �`�[�����ƂɎg�p����ϐ�
    Dim teamID As Integer
    Dim teamName As String
    Dim numberOfGamesPlayedBeforeThisSection, numberOfGamesPlayedAfterThisSection As Integer
    Dim pcrPositiveRate, accidentBonusValue As Single
    Dim pcrResultText, pcrResultMessage As String
    
    ' �I�育�ƂɎg�p����ϐ�
    Dim playerID, rowIdx As Integer
    Dim playerName, playerNameRegistered As String
    Dim accidentRate As Single
    Dim accidentPeriod As Integer
    Dim accidentPeriodString, accidentOverview, accidentText, accidentMessage As String
    
    ' ���̑��ϐ�
    Dim columnIdxOfStamina, columnIdx As Integer
    Dim dice As Single
    
    Randomize
    
    ' �V�[�Y���A�ߐi�s�󋵂̊m�F
    seasonName = ActiveSheet.Cells(1, "A").Value
    numberOfSection = WorksheetFunction.CountIf(ActiveSheet.Range("BA2:BA241"), 0) / 8
    If numberOfSection = 0 Then
        allAccidentResultNotification = "�y�J���O�X�y����z"
    Else
        allAccidentResultNotification = "�y��" & numberOfSection & "�ߏI�����X�y����z"
    End If
    rankNotification = "�yMPB�j���[�X�z"
    If numberOfSection = 25 Then
        rankNotification = rankNotification & vbCrLf & _
                           "�E<�d�v>�������AB9GG�m�~�l�[�g�I�[�_�[�̒�o��t���J�n���܂��B�Y�؂͑�28�ߏI�����ł��B�e�`�[���ALINE�O���[�v�̃A���o���ɒ�o�����肢�������܂��B"
    ElseIf numberOfSection = 26 Or numberOfSection = 27 Then
        rankNotification = rankNotification & vbCrLf & _
                           "�E<�d�v>B9GG�m�~�l�[�g�I�[�_�[���o��t���ł��B����o�̃`�[���́A��28�߂��I������܂łɁALINE�O���[�v�̃A���o���ւ̒�o�����肢�������܂��B"
    ElseIf numberOfSection = 28 Then
        rankNotification = rankNotification & vbCrLf & _
                           "�E<�d�v>B9GG�m�~�l�[�g�I�[�_�[�̒�o/�ύX��t����ߐ؂�܂����B"
    End If
    tsobChangeNotification = "�yTS/OB�g�U�蒼���̂��m�点�z"
    
    If Not debugModeFlg Then
        ' ���߂̖��J�n�Ɨ\���攭�̏o�������m�F
        If numberOfSection > 0 Then
            If ActiveSheet.Cells(numberOfSection * 8 + 2, "D").Value = "" Or ActiveSheet.Cells(numberOfSection * 8 + 6, "D").Value = "" Or _
               ActiveSheet.Cells(numberOfSection * 8 + 2, "H").Value = "" Or ActiveSheet.Cells(numberOfSection * 8 + 6, "H").Value = "" Then
                MsgBox "��" & numberOfSection + 1 & "�߂̐攭�\�����������Ă��܂���B"
                End
            End If
        End If
        
        If ActiveSheet.Cells(numberOfSection * 8 + 3, "D").Value <> "" Or ActiveSheet.Cells(numberOfSection * 8 + 7, "D").Value <> "" Or _
           ActiveSheet.Cells(numberOfSection * 8 + 3, "F").Value <> "" Or ActiveSheet.Cells(numberOfSection * 8 + 7, "F").Value <> "" Or _
           ActiveSheet.Cells(numberOfSection * 8 + 3, "H").Value <> "" Or ActiveSheet.Cells(numberOfSection * 8 + 7, "H").Value <> "" Then
            MsgBox "��" & numberOfSection + 1 & "�߂̎������ʂ��s���ɓ��͂���Ă��܂��B"
            End
        End If
    End If
    
    ' �����ɒǋL
    ' ��O�I�ɂ����ŕϐ��錾���s��
    Dim tmp1, tmp2 As String
    Dim vStatus(6, 6) As String ' [0:���ߎ��{�O,1:���ߎ��{��,2:���߁�-��/��-��,3:���߁�-��/��-��,4:���߁�-��/��-��5:���߁�-��/��-��][teamID+flag]��{�D��,M*,����V����,-}
    Dim teamNameOfNextSection(2, 2) As String ' [0:�@,1:�A][0:Home,1:Visitor]
    
    If numberOfSection > 0 And numberOfSection < 30 Then
        ' ���ߎ��{�O�̏󋵊m�F
        tmp1 = ActiveSheet.Cells(numberOfSection * 8 - 5, "F").Value
        tmp2 = ActiveSheet.Cells(numberOfSection * 8 - 1, "F").Value
        ActiveSheet.Cells(numberOfSection * 8 - 5, "F").Value = ""
        ActiveSheet.Cells(numberOfSection * 8 - 1, "F").Value = ""
        Application.Calculate
        
        vStatus(0, 5) = "false"
        
        For teamID = 0 To 4
            
            vStatus(0, teamID) = "-"
            
            If Sheets(seasonName & "_�e��L�^").Cells(teamID + 2, "BR").Value = 0 Then
                vStatus(0, teamID) = "����V����"
            ElseIf Sheets(seasonName & "_�e��L�^").Cells(teamID + 2, "BX").Value <> "-" Then
                vStatus(0, teamID) = Sheets(seasonName & "_�e��L�^").Cells(teamID + 2, "BX").Value
            End If
            
        Next teamID
        
        ' ���ߎ��{��̏󋵊m�F
        ActiveSheet.Cells(numberOfSection * 8 - 5, "F").Value = tmp1
        ActiveSheet.Cells(numberOfSection * 8 - 1, "F").Value = tmp2
        Application.Calculate
        
        vStatus(1, 5) = "false"
        
        For teamID = 0 To 4
            
            vStatus(1, teamID) = "-"
            
            If Sheets(seasonName & "_�e��L�^").Cells(teamID + 2, "BR").Value = 0 Then
                vStatus(1, teamID) = "����V����"
            ElseIf Sheets(seasonName & "_�e��L�^").Cells(teamID + 2, "BX").Value <> "-" Then
                vStatus(1, teamID) = Sheets(seasonName & "_�e��L�^").Cells(teamID + 2, "BX").Value
            End If
            
            If Left(vStatus(0, teamID), 1) <> Left(vStatus(1, teamID), 1) Then
                vStatus(1, 5) = "true"
            End If
            
        Next teamID
        
        ' ���߁�-��/��-���̏󋵊m�F
        ActiveSheet.Cells(numberOfSection * 8 + 3, "D").Value = "X"
        ActiveSheet.Cells(numberOfSection * 8 + 3, "F").Value = "-"
        ActiveSheet.Cells(numberOfSection * 8 + 3, "H").Value = "0"
        ActiveSheet.Cells(numberOfSection * 8 + 7, "D").Value = "X"
        ActiveSheet.Cells(numberOfSection * 8 + 7, "F").Value = "-"
        ActiveSheet.Cells(numberOfSection * 8 + 7, "H").Value = "0"
        Application.Calculate
        
        vStatus(2, 5) = "false"
        
        For teamID = 0 To 4
            
            vStatus(2, teamID) = "-"
            
            If Sheets(seasonName & "_�e��L�^").Cells(teamID + 2, "BR").Value = 0 Then
                vStatus(2, teamID) = "����V����"
            ElseIf Sheets(seasonName & "_�e��L�^").Cells(teamID + 2, "BX").Value <> "-" Then
                vStatus(2, teamID) = Sheets(seasonName & "_�e��L�^").Cells(teamID + 2, "BX").Value
            End If
            
            If Left(vStatus(1, teamID), 1) <> Left(vStatus(2, teamID), 1) Then
                vStatus(2, 5) = "true"
            End If
            
        Next teamID
        
        ' ���߁�-��/��-���̏󋵊m�F
        ActiveSheet.Cells(numberOfSection * 8 + 3, "D").Value = "X"
        ActiveSheet.Cells(numberOfSection * 8 + 3, "F").Value = "-"
        ActiveSheet.Cells(numberOfSection * 8 + 3, "H").Value = "0"
        ActiveSheet.Cells(numberOfSection * 8 + 7, "D").Value = "0"
        ActiveSheet.Cells(numberOfSection * 8 + 7, "F").Value = "-"
        ActiveSheet.Cells(numberOfSection * 8 + 7, "H").Value = "X"
        Application.Calculate
        
        vStatus(3, 5) = "false"
        
        For teamID = 0 To 4
            
            vStatus(3, teamID) = "-"
            
            If Sheets(seasonName & "_�e��L�^").Cells(teamID + 2, "BR").Value = 0 Then
                vStatus(3, teamID) = "����V����"
            ElseIf Sheets(seasonName & "_�e��L�^").Cells(teamID + 2, "BX").Value <> "-" Then
                vStatus(3, teamID) = Sheets(seasonName & "_�e��L�^").Cells(teamID + 2, "BX").Value
            End If
            
            If Left(vStatus(1, teamID), 1) <> Left(vStatus(3, teamID), 1) Then
                vStatus(3, 5) = "true"
            End If
            
        Next teamID
        
        ' ���߁�-��/��-���̏󋵊m�F
        ActiveSheet.Cells(numberOfSection * 8 + 3, "D").Value = "0"
        ActiveSheet.Cells(numberOfSection * 8 + 3, "F").Value = "-"
        ActiveSheet.Cells(numberOfSection * 8 + 3, "H").Value = "X"
        ActiveSheet.Cells(numberOfSection * 8 + 7, "D").Value = "X"
        ActiveSheet.Cells(numberOfSection * 8 + 7, "F").Value = "-"
        ActiveSheet.Cells(numberOfSection * 8 + 7, "H").Value = "0"
        Application.Calculate
        
        vStatus(4, 5) = "false"
        
        For teamID = 0 To 4
            
            vStatus(4, teamID) = "-"
            
            If Sheets(seasonName & "_�e��L�^").Cells(teamID + 2, "BR").Value = 0 Then
                vStatus(4, teamID) = "����V����"
            ElseIf Sheets(seasonName & "_�e��L�^").Cells(teamID + 2, "BX").Value <> "-" Then
                vStatus(4, teamID) = Sheets(seasonName & "_�e��L�^").Cells(teamID + 2, "BX").Value
            End If
            
            If Left(vStatus(1, teamID), 1) <> Left(vStatus(4, teamID), 1) Then
                vStatus(4, 5) = "true"
            End If
            
        Next teamID
        
        ' ���߁�-��/��-���̏󋵊m�F
        ActiveSheet.Cells(numberOfSection * 8 + 3, "D").Value = "0"
        ActiveSheet.Cells(numberOfSection * 8 + 3, "F").Value = "-"
        ActiveSheet.Cells(numberOfSection * 8 + 3, "H").Value = "X"
        ActiveSheet.Cells(numberOfSection * 8 + 7, "D").Value = "0"
        ActiveSheet.Cells(numberOfSection * 8 + 7, "F").Value = "-"
        ActiveSheet.Cells(numberOfSection * 8 + 7, "H").Value = "X"
        Application.Calculate
        
        vStatus(5, 5) = "false"
        
        For teamID = 0 To 4
            
            vStatus(5, teamID) = "-"
            
            If Sheets(seasonName & "_�e��L�^").Cells(teamID + 2, "BR").Value = 0 Then
                vStatus(5, teamID) = "����V����"
            ElseIf Sheets(seasonName & "_�e��L�^").Cells(teamID + 2, "BX").Value <> "-" Then
                vStatus(5, teamID) = Sheets(seasonName & "_�e��L�^").Cells(teamID + 2, "BX").Value
            End If
            
            If Left(vStatus(1, teamID), 1) <> Left(vStatus(5, teamID), 1) Then
                vStatus(5, 5) = "true"
            End If
            
        Next teamID
        
        ActiveSheet.Cells(numberOfSection * 8 + 3, "D").Value = ""
        ActiveSheet.Cells(numberOfSection * 8 + 3, "F").Value = ""
        ActiveSheet.Cells(numberOfSection * 8 + 3, "H").Value = ""
        ActiveSheet.Cells(numberOfSection * 8 + 7, "D").Value = ""
        ActiveSheet.Cells(numberOfSection * 8 + 7, "F").Value = ""
        ActiveSheet.Cells(numberOfSection * 8 + 7, "H").Value = ""
        Application.Calculate
        
        ' ���߂܂łɗD�������܂��Ă���ꍇ�X�L�b�v
        For teamID = 0 To 4
            
            If vStatus(0, teamID) = "�D��" Then
                GoTo MPB_NEWS_CHECK_END_POINT
            ElseIf vStatus(1, teamID) = "�D��" Then
                rankNotification = rankNotification & vbCrLf & _
                                   "�E" & Sheets(seasonName & "_�e��L�^").Cells(teamID + 2, "BA").Value & ":���[�O�D�����m��I"
                GoTo MPB_NEWS_CHECK_END_POINT
            End If
            
        Next teamID
        
        If vStatus(1, 5) = "true" Then
            ' ���߂̃}�W�b�N����
            For teamID = 0 To 4
                
                If Left(vStatus(0, teamID), 1) = "M" And Left(vStatus(1, teamID), 1) <> "M" Then
                    rankNotification = rankNotification & vbCrLf & _
                                       "�E" & Sheets(seasonName & "_�e��L�^").Cells(teamID + 2, "BA").Value & ":�D���}�W�b�N�����Łc"
                End If
                
            Next teamID
            
            ' ���߂̎���V����
            For teamID = 0 To 4
                
                If vStatus(0, teamID) <> "����V����" And vStatus(1, teamID) = "����V����" Then
                    rankNotification = rankNotification & vbCrLf & _
                                       "�E" & Sheets(seasonName & "_�e��L�^").Cells(teamID + 2, "BA").Value & ":���͗D�������Łc"
                End If
                
            Next teamID
            
            ' ���߂̎���V����
            For teamID = 0 To 4
                
                If vStatus(0, teamID) = "����V����" And vStatus(1, teamID) <> "����V����" Then
                    rankNotification = rankNotification & vbCrLf & _
                                       "�E" & Sheets(seasonName & "_�e��L�^").Cells(teamID + 2, "BA").Value & ":���͗D���������I"
                End If
                
            Next teamID
            
            ' ���߂̃}�W�b�N�_��
            For teamID = 0 To 4
                
                If Left(vStatus(0, teamID), 1) <> "M" And Left(vStatus(1, teamID), 1) = "M" Then
                    rankNotification = rankNotification & vbCrLf & _
                                       "�E" & Sheets(seasonName & "_�e��L�^").Cells(teamID + 2, "BA").Value & ":�D���}�W�b�N(" & vStatus(1, teamID) & ")���_���I"
                End If
                
            Next teamID
            
        End If
        
        ' ���ߎ��{��̓W�]
        teamNameOfNextSection(0, 0) = ActiveSheet.Cells(numberOfSection * 8 + 2, "C").Value
        teamNameOfNextSection(0, 1) = ActiveSheet.Cells(numberOfSection * 8 + 2, "J").Value
        teamNameOfNextSection(1, 0) = ActiveSheet.Cells(numberOfSection * 8 + 6, "C").Value
        teamNameOfNextSection(1, 1) = ActiveSheet.Cells(numberOfSection * 8 + 6, "J").Value
        
        ' �@��-���ŋ���
        If vStatus(2, 5) = "true" And vStatus(3, 5) = "true" Then
            ' �D��
            For teamID = 0 To 4
                
                If Left(vStatus(1, teamID), 1) <> "�D" And Left(vStatus(2, teamID), 1) = "�D" And Left(vStatus(3, teamID), 1) = "�D" Then
                    rankNotification = rankNotification & vbCrLf & _
                                       "�E���� " & teamNameOfNextSection(0, 0) & "��-��" & teamNameOfNextSection(0, 1) & " �ŁA" & vbCrLf & _
                                       "�@" & Sheets(seasonName & "_�e��L�^").Cells(teamID + 2, "BA").Value & "�̃��[�O�D�����m��I"
                    GoTo MPB_NEWS_PATTERN1_END_POINT
                End If
                
            Next teamID
            
            ' �}�W�b�N����
            For teamID = 0 To 4
                
                If Left(vStatus(1, teamID), 1) = "M" And Left(vStatus(2, teamID), 1) <> "M" And Left(vStatus(3, teamID), 1) <> "M" Then
                    rankNotification = rankNotification & vbCrLf & _
                                       "�E���� " & teamNameOfNextSection(0, 0) & "��-��" & teamNameOfNextSection(0, 1) & " �ŁA" & vbCrLf & _
                                       "�@" & Sheets(seasonName & "_�e��L�^").Cells(teamID + 2, "BA").Value & "�̗D���}�W�b�N�����Łc"
                End If
                
            Next teamID
            
            ' ����V����
            For teamID = 0 To 4
                
                If Left(vStatus(1, teamID), 1) <> "��" And Left(vStatus(2, teamID), 1) = "��" And Left(vStatus(3, teamID), 1) = "��" Then
                    rankNotification = rankNotification & vbCrLf & _
                                       "�E���� " & teamNameOfNextSection(0, 0) & "��-��" & teamNameOfNextSection(0, 1) & " �ŁA" & vbCrLf & _
                                       "�@" & Sheets(seasonName & "_�e��L�^").Cells(teamID + 2, "BA").Value & "�̎��͗D�������Łc"
                End If
                
            Next teamID
            
            ' ����V����
            For teamID = 0 To 4
                
                If Left(vStatus(1, teamID), 1) = "��" And Left(vStatus(2, teamID), 1) <> "��" And Left(vStatus(3, teamID), 1) <> "��" Then
                    rankNotification = rankNotification & vbCrLf & _
                                       "�E���� " & teamNameOfNextSection(0, 0) & "��-��" & teamNameOfNextSection(0, 1) & " �ŁA" & vbCrLf & _
                                       "�@" & Sheets(seasonName & "_�e��L�^").Cells(teamID + 2, "BA").Value & "�̎��͗D���������I"
                End If
                
            Next teamID
            
            ' �}�W�b�N�_��
            For teamID = 0 To 4
                
                If Left(vStatus(1, teamID), 1) <> "M" And Left(vStatus(2, teamID), 1) = "M" And Left(vStatus(3, teamID), 1) = "M" Then
                    rankNotification = rankNotification & vbCrLf & _
                                       "�E���� " & teamNameOfNextSection(0, 0) & "��-��" & teamNameOfNextSection(0, 1) & " �ŁA" & vbCrLf & _
                                       "�@" & Sheets(seasonName & "_�e��L�^").Cells(teamID + 2, "BA").Value & "�̗D���}�W�b�N���_���I"
                End If
                
            Next teamID
            
        End If
        
MPB_NEWS_PATTERN1_END_POINT:

        ' �@��-���ŋ���
        If vStatus(4, 5) = "true" And vStatus(5, 5) = "true" Then
            ' �D��
            For teamID = 0 To 4
                
                If Left(vStatus(1, teamID), 1) <> "�D" And Left(vStatus(4, teamID), 1) = "�D" And Left(vStatus(5, teamID), 1) = "�D" Then
                    rankNotification = rankNotification & vbCrLf & _
                                       "�E���� " & teamNameOfNextSection(0, 0) & "��-��" & teamNameOfNextSection(0, 1) & " �ŁA" & vbCrLf & _
                                       "�@" & Sheets(seasonName & "_�e��L�^").Cells(teamID + 2, "BA").Value & "�̃��[�O�D�����m��I"
                    GoTo MPB_NEWS_PATTERN2_END_POINT
                End If
                
            Next teamID
            
            ' �}�W�b�N����
            For teamID = 0 To 4
                
                If Left(vStatus(1, teamID), 1) = "M" And Left(vStatus(4, teamID), 1) <> "M" And Left(vStatus(5, teamID), 1) <> "M" Then
                    rankNotification = rankNotification & vbCrLf & _
                                       "�E���� " & teamNameOfNextSection(0, 0) & "��-��" & teamNameOfNextSection(0, 1) & " �ŁA" & vbCrLf & _
                                       "�@" & Sheets(seasonName & "_�e��L�^").Cells(teamID + 2, "BA").Value & "�̗D���}�W�b�N�����Łc"
                End If
                
            Next teamID
            
            ' ����V����
            For teamID = 0 To 4
                
                If Left(vStatus(1, teamID), 1) <> "��" And Left(vStatus(4, teamID), 1) = "��" And Left(vStatus(5, teamID), 1) = "��" Then
                    rankNotification = rankNotification & vbCrLf & _
                                       "�E���� " & teamNameOfNextSection(0, 0) & "��-��" & teamNameOfNextSection(0, 1) & " �ŁA" & vbCrLf & _
                                       "�@" & Sheets(seasonName & "_�e��L�^").Cells(teamID + 2, "BA").Value & "�̎��͗D�������Łc"
                End If
                
            Next teamID
            
            ' ����V����
            For teamID = 0 To 4
                
                If Left(vStatus(1, teamID), 1) = "��" And Left(vStatus(4, teamID), 1) <> "��" And Left(vStatus(5, teamID), 1) <> "��" Then
                    rankNotification = rankNotification & vbCrLf & _
                                       "�E���� " & teamNameOfNextSection(0, 0) & "��-��" & teamNameOfNextSection(0, 1) & " �ŁA" & vbCrLf & _
                                       "�@" & Sheets(seasonName & "_�e��L�^").Cells(teamID + 2, "BA").Value & "�̎��͗D���������I"
                End If
                
            Next teamID
            
            ' �}�W�b�N�_��
            For teamID = 0 To 4
                
                If Left(vStatus(1, teamID), 1) <> "M" And Left(vStatus(4, teamID), 1) = "M" And Left(vStatus(5, teamID), 1) = "M" Then
                    rankNotification = rankNotification & vbCrLf & _
                                       "�E���� " & teamNameOfNextSection(0, 0) & "��-��" & teamNameOfNextSection(0, 1) & " �ŁA" & vbCrLf & _
                                       "�@" & Sheets(seasonName & "_�e��L�^").Cells(teamID + 2, "BA").Value & "�̗D���}�W�b�N���_���I"
                End If
                
            Next teamID
            
        End If
        
MPB_NEWS_PATTERN2_END_POINT:

        ' �A��-���ŋ���
        If vStatus(2, 5) = "true" And vStatus(4, 5) = "true" Then
            ' �D��
            For teamID = 0 To 4
                
                If Left(vStatus(1, teamID), 1) <> "�D" And Left(vStatus(2, teamID), 1) = "�D" And Left(vStatus(4, teamID), 1) = "�D" Then
                    rankNotification = rankNotification & vbCrLf & _
                                       "�E���� " & teamNameOfNextSection(1, 0) & "��-��" & teamNameOfNextSection(1, 1) & " �ŁA" & vbCrLf & _
                                       "�@" & Sheets(seasonName & "_�e��L�^").Cells(teamID + 2, "BA").Value & "�̃��[�O�D�����m��I"
                    GoTo MPB_NEWS_PATTERN3_END_POINT
                End If
                
            Next teamID
            
            ' �}�W�b�N����
            For teamID = 0 To 4
                
                If Left(vStatus(1, teamID), 1) = "M" And Left(vStatus(2, teamID), 1) <> "M" And Left(vStatus(4, teamID), 1) <> "M" Then
                    rankNotification = rankNotification & vbCrLf & _
                                       "�E���� " & teamNameOfNextSection(1, 0) & "��-��" & teamNameOfNextSection(1, 1) & " �ŁA" & vbCrLf & _
                                       "�@" & Sheets(seasonName & "_�e��L�^").Cells(teamID + 2, "BA").Value & "�̗D���}�W�b�N�����Łc"
                End If
                
            Next teamID
            
            ' ����V����
            For teamID = 0 To 4
                
                If Left(vStatus(1, teamID), 1) <> "��" And Left(vStatus(2, teamID), 1) = "��" And Left(vStatus(4, teamID), 1) = "��" Then
                    rankNotification = rankNotification & vbCrLf & _
                                       "�E���� " & teamNameOfNextSection(1, 0) & "��-��" & teamNameOfNextSection(1, 1) & " �ŁA" & vbCrLf & _
                                       "�@" & Sheets(seasonName & "_�e��L�^").Cells(teamID + 2, "BA").Value & "�̎��͗D�������Łc"
                End If
                
            Next teamID
            
            ' ����V����
            For teamID = 0 To 4
                
                If Left(vStatus(1, teamID), 1) = "��" And Left(vStatus(2, teamID), 1) <> "��" And Left(vStatus(4, teamID), 1) <> "��" Then
                    rankNotification = rankNotification & vbCrLf & _
                                       "�E���� " & teamNameOfNextSection(1, 0) & "��-��" & teamNameOfNextSection(1, 1) & " �ŁA" & vbCrLf & _
                                       "�@" & Sheets(seasonName & "_�e��L�^").Cells(teamID + 2, "BA").Value & "�̎��͗D���������I"
                End If
                
            Next teamID
            
            ' �}�W�b�N�_��
            For teamID = 0 To 4
                
                If Left(vStatus(1, teamID), 1) <> "M" And Left(vStatus(2, teamID), 1) = "M" And Left(vStatus(4, teamID), 1) = "M" Then
                    rankNotification = rankNotification & vbCrLf & _
                                       "�E���� " & teamNameOfNextSection(1, 0) & "��-��" & teamNameOfNextSection(1, 1) & " �ŁA" & vbCrLf & _
                                       "�@" & Sheets(seasonName & "_�e��L�^").Cells(teamID + 2, "BA").Value & "�̗D���}�W�b�N���_���I"
                End If
                
            Next teamID
            
        End If
        
MPB_NEWS_PATTERN3_END_POINT:

        ' �A��-���ŋ���
        If vStatus(3, 5) = "true" And vStatus(5, 5) = "true" Then
            ' �D��
            For teamID = 0 To 4
                
                If Left(vStatus(1, teamID), 1) <> "�D" And Left(vStatus(3, teamID), 1) = "�D" And Left(vStatus(5, teamID), 1) = "�D" Then
                    rankNotification = rankNotification & vbCrLf & _
                                       "�E���� " & teamNameOfNextSection(1, 0) & "��-��" & teamNameOfNextSection(1, 1) & " �ŁA" & vbCrLf & _
                                       "�@" & Sheets(seasonName & "_�e��L�^").Cells(teamID + 2, "BA").Value & "�̃��[�O�D�����m��I"
                    GoTo MPB_NEWS_PATTERN4_END_POINT
                End If
                
            Next teamID
            
            ' �}�W�b�N����
            For teamID = 0 To 4
                
                If Left(vStatus(1, teamID), 1) = "M" And Left(vStatus(3, teamID), 1) <> "M" And Left(vStatus(5, teamID), 1) <> "M" Then
                    rankNotification = rankNotification & vbCrLf & _
                                       "�E���� " & teamNameOfNextSection(1, 0) & "��-��" & teamNameOfNextSection(1, 1) & " �ŁA" & vbCrLf & _
                                       "�@" & Sheets(seasonName & "_�e��L�^").Cells(teamID + 2, "BA").Value & "�̗D���}�W�b�N�����Łc"
                End If
                
            Next teamID
            
            ' ����V����
            For teamID = 0 To 4
                
                If Left(vStatus(1, teamID), 1) <> "��" And Left(vStatus(3, teamID), 1) = "��" And Left(vStatus(5, teamID), 1) = "��" Then
                    rankNotification = rankNotification & vbCrLf & _
                                       "�E���� " & teamNameOfNextSection(1, 0) & "��-��" & teamNameOfNextSection(1, 1) & " �ŁA" & vbCrLf & _
                                       "�@" & Sheets(seasonName & "_�e��L�^").Cells(teamID + 2, "BA").Value & "�̎��͗D�������Łc"
                End If
                
            Next teamID
            
            ' ����V����
            For teamID = 0 To 4
                
                If Left(vStatus(1, teamID), 1) = "��" And Left(vStatus(3, teamID), 1) <> "��" And Left(vStatus(5, teamID), 1) <> "��" Then
                    rankNotification = rankNotification & vbCrLf & _
                                       "�E���� " & teamNameOfNextSection(1, 0) & "��-��" & teamNameOfNextSection(1, 1) & " �ŁA" & vbCrLf & _
                                       "�@" & Sheets(seasonName & "_�e��L�^").Cells(teamID + 2, "BA").Value & "�̎��͗D���������I"
                End If
                
            Next teamID
            
            ' �}�W�b�N�_��
            For teamID = 0 To 4
                
                If Left(vStatus(1, teamID), 1) <> "M" And Left(vStatus(3, teamID), 1) = "M" And Left(vStatus(5, teamID), 1) = "M" Then
                    rankNotification = rankNotification & vbCrLf & _
                                       "�E���� " & teamNameOfNextSection(1, 0) & "��-��" & teamNameOfNextSection(1, 1) & " �ŁA" & vbCrLf & _
                                       "�@" & Sheets(seasonName & "_�e��L�^").Cells(teamID + 2, "BA").Value & "�̗D���}�W�b�N���_���I"
                End If
                
            Next teamID
            
        End If
        
MPB_NEWS_PATTERN4_END_POINT:

        ' �@��-���A��-��
        If vStatus(2, 5) = "true" Then
            ' �D��
            For teamID = 0 To 4
                
                If Left(vStatus(1, teamID), 1) <> "�D" And Left(vStatus(2, teamID), 1) = "�D" And Left(vStatus(3, teamID), 1) <> "�D" And Left(vStatus(4, teamID), 1) <> "�D" Then
                    rankNotification = rankNotification & vbCrLf & _
                                       "�E���� " & teamNameOfNextSection(0, 0) & "��-��" & teamNameOfNextSection(0, 1) & " & " & teamNameOfNextSection(1, 0) & "��-��" & teamNameOfNextSection(1, 1) & " �ŁA" & vbCrLf & _
                                       "�@" & Sheets(seasonName & "_�e��L�^").Cells(teamID + 2, "BA").Value & "�̃��[�O�D�����m��I"
                    GoTo MPB_NEWS_PATTERN5_END_POINT
                End If
                
            Next teamID
            
            ' �}�W�b�N����
            For teamID = 0 To 4
                
                If Left(vStatus(1, teamID), 1) = "M" And Left(vStatus(2, teamID), 1) <> "M" And Left(vStatus(3, teamID), 1) = "M" And Left(vStatus(4, teamID), 1) = "M" Then
                    rankNotification = rankNotification & vbCrLf & _
                                       "�E���� " & teamNameOfNextSection(0, 0) & "��-��" & teamNameOfNextSection(0, 1) & " & " & teamNameOfNextSection(1, 0) & "��-��" & teamNameOfNextSection(1, 1) & " �ŁA" & vbCrLf & _
                                       "�@" & Sheets(seasonName & "_�e��L�^").Cells(teamID + 2, "BA").Value & "�̗D���}�W�b�N�����Łc"
                End If
                
            Next teamID
            
            ' ����V����
            For teamID = 0 To 4
                
                If Left(vStatus(1, teamID), 1) <> "��" And Left(vStatus(2, teamID), 1) = "��" And Left(vStatus(3, teamID), 1) <> "��" And Left(vStatus(4, teamID), 1) <> "��" Then
                    rankNotification = rankNotification & vbCrLf & _
                                       "�E���� " & teamNameOfNextSection(0, 0) & "��-��" & teamNameOfNextSection(0, 1) & " & " & teamNameOfNextSection(1, 0) & "��-��" & teamNameOfNextSection(1, 1) & " �ŁA" & vbCrLf & _
                                       "�@" & Sheets(seasonName & "_�e��L�^").Cells(teamID + 2, "BA").Value & "�̎��͗D�������Łc"
                End If
                
            Next teamID
            
            ' ����V����
            For teamID = 0 To 4
                
                If Left(vStatus(1, teamID), 1) = "��" And Left(vStatus(2, teamID), 1) <> "��" And Left(vStatus(3, teamID), 1) = "��" And Left(vStatus(4, teamID), 1) = "��" Then
                    rankNotification = rankNotification & vbCrLf & _
                                       "�E���� " & teamNameOfNextSection(0, 0) & "��-��" & teamNameOfNextSection(0, 1) & " & " & teamNameOfNextSection(1, 0) & "��-��" & teamNameOfNextSection(1, 1) & " �ŁA" & vbCrLf & _
                                       "�@" & Sheets(seasonName & "_�e��L�^").Cells(teamID + 2, "BA").Value & "�̎��͗D���������I"
                End If
                
            Next teamID
            
            ' �}�W�b�N�_��
            For teamID = 0 To 4
                
                If Left(vStatus(1, teamID), 1) <> "M" And Left(vStatus(2, teamID), 1) = "M" And Left(vStatus(3, teamID), 1) <> "M" And Left(vStatus(4, teamID), 1) <> "M" Then
                    rankNotification = rankNotification & vbCrLf & _
                                       "�E���� " & teamNameOfNextSection(0, 0) & "��-��" & teamNameOfNextSection(0, 1) & " & " & teamNameOfNextSection(1, 0) & "��-��" & teamNameOfNextSection(1, 1) & " �ŁA" & vbCrLf & _
                                       "�@" & Sheets(seasonName & "_�e��L�^").Cells(teamID + 2, "BA").Value & "�̗D���}�W�b�N���_���I"
                End If
                
            Next teamID
            
        End If
        
MPB_NEWS_PATTERN5_END_POINT:

        ' �@��-���A��-��
        If vStatus(3, 5) = "true" Then
            ' �D��
            For teamID = 0 To 4
                
                If Left(vStatus(1, teamID), 1) <> "�D" And Left(vStatus(2, teamID), 1) <> "�D" And Left(vStatus(3, teamID), 1) = "�D" And Left(vStatus(5, teamID), 1) <> "�D" Then
                    rankNotification = rankNotification & vbCrLf & _
                                       "�E���� " & teamNameOfNextSection(0, 0) & "��-��" & teamNameOfNextSection(0, 1) & " & " & teamNameOfNextSection(1, 0) & "��-��" & teamNameOfNextSection(1, 1) & " �ŁA" & vbCrLf & _
                                       "�@" & Sheets(seasonName & "_�e��L�^").Cells(teamID + 2, "BA").Value & "�̃��[�O�D�����m��I"
                    GoTo MPB_NEWS_PATTERN6_END_POINT
                End If
                
            Next teamID
            
            ' �}�W�b�N����
            For teamID = 0 To 4
                
                If Left(vStatus(1, teamID), 1) = "M" And Left(vStatus(2, teamID), 1) = "M" And Left(vStatus(3, teamID), 1) <> "M" And Left(vStatus(5, teamID), 1) = "M" Then
                    rankNotification = rankNotification & vbCrLf & _
                                       "�E���� " & teamNameOfNextSection(0, 0) & "��-��" & teamNameOfNextSection(0, 1) & " & " & teamNameOfNextSection(1, 0) & "��-��" & teamNameOfNextSection(1, 1) & " �ŁA" & vbCrLf & _
                                       "�@" & Sheets(seasonName & "_�e��L�^").Cells(teamID + 2, "BA").Value & "�̗D���}�W�b�N�����Łc"
                End If
                
            Next teamID
            
            ' ����V����
            For teamID = 0 To 4
                
                If Left(vStatus(1, teamID), 1) <> "��" And Left(vStatus(2, teamID), 1) <> "��" And Left(vStatus(3, teamID), 1) = "��" And Left(vStatus(5, teamID), 1) <> "��" Then
                    rankNotification = rankNotification & vbCrLf & _
                                       "�E���� " & teamNameOfNextSection(0, 0) & "��-��" & teamNameOfNextSection(0, 1) & " & " & teamNameOfNextSection(1, 0) & "��-��" & teamNameOfNextSection(1, 1) & " �ŁA" & vbCrLf & _
                                       "�@" & Sheets(seasonName & "_�e��L�^").Cells(teamID + 2, "BA").Value & "�̎��͗D�������Łc"
                End If
                
            Next teamID
            
            ' ����V����
            For teamID = 0 To 4
                
                If Left(vStatus(1, teamID), 1) = "��" And Left(vStatus(2, teamID), 1) = "��" And Left(vStatus(3, teamID), 1) <> "��" And Left(vStatus(5, teamID), 1) = "��" Then
                    rankNotification = rankNotification & vbCrLf & _
                                       "�E���� " & teamNameOfNextSection(0, 0) & "��-��" & teamNameOfNextSection(0, 1) & " & " & teamNameOfNextSection(1, 0) & "��-��" & teamNameOfNextSection(1, 1) & " �ŁA" & vbCrLf & _
                                       "�@" & Sheets(seasonName & "_�e��L�^").Cells(teamID + 2, "BA").Value & "�̎��͗D���������I"
                End If
                
            Next teamID
            
            ' �}�W�b�N�_��
            For teamID = 0 To 4
                
                If Left(vStatus(1, teamID), 1) <> "M" And Left(vStatus(2, teamID), 1) <> "M" And Left(vStatus(3, teamID), 1) = "M" And Left(vStatus(5, teamID), 1) <> "M" Then
                    rankNotification = rankNotification & vbCrLf & _
                                       "�E���� " & teamNameOfNextSection(0, 0) & "��-��" & teamNameOfNextSection(0, 1) & " & " & teamNameOfNextSection(1, 0) & "��-��" & teamNameOfNextSection(1, 1) & " �ŁA" & vbCrLf & _
                                       "�@" & Sheets(seasonName & "_�e��L�^").Cells(teamID + 2, "BA").Value & "�̗D���}�W�b�N���_���I"
                End If
                
            Next teamID
            
        End If
        
MPB_NEWS_PATTERN6_END_POINT:

        ' �@��-���A��-��
        If vStatus(4, 5) = "true" Then
            ' �D��
            For teamID = 0 To 4
                
                If Left(vStatus(1, teamID), 1) <> "�D" And Left(vStatus(2, teamID), 1) <> "�D" And Left(vStatus(4, teamID), 1) = "�D" And Left(vStatus(5, teamID), 1) <> "�D" Then
                    rankNotification = rankNotification & vbCrLf & _
                                       "�E���� " & teamNameOfNextSection(0, 0) & "��-��" & teamNameOfNextSection(0, 1) & " & " & teamNameOfNextSection(1, 0) & "��-��" & teamNameOfNextSection(1, 1) & " �ŁA" & vbCrLf & _
                                       "�@" & Sheets(seasonName & "_�e��L�^").Cells(teamID + 2, "BA").Value & "�̃��[�O�D�����m��I"
                    GoTo MPB_NEWS_PATTERN7_END_POINT
                End If
                
            Next teamID
            
            ' �}�W�b�N����
            For teamID = 0 To 4
                
                If Left(vStatus(1, teamID), 1) = "M" And Left(vStatus(2, teamID), 1) = "M" And Left(vStatus(4, teamID), 1) <> "M" And Left(vStatus(5, teamID), 1) = "M" Then
                    rankNotification = rankNotification & vbCrLf & _
                                       "�E���� " & teamNameOfNextSection(0, 0) & "��-��" & teamNameOfNextSection(0, 1) & " & " & teamNameOfNextSection(1, 0) & "��-��" & teamNameOfNextSection(1, 1) & " �ŁA" & vbCrLf & _
                                       "�@" & Sheets(seasonName & "_�e��L�^").Cells(teamID + 2, "BA").Value & "�̗D���}�W�b�N�����Łc"
                End If
                
            Next teamID
            
            ' ����V����
            For teamID = 0 To 4
                
                If Left(vStatus(1, teamID), 1) <> "��" And Left(vStatus(2, teamID), 1) <> "��" And Left(vStatus(4, teamID), 1) = "��" And Left(vStatus(5, teamID), 1) <> "��" Then
                    rankNotification = rankNotification & vbCrLf & _
                                       "�E���� " & teamNameOfNextSection(0, 0) & "��-��" & teamNameOfNextSection(0, 1) & " & " & teamNameOfNextSection(1, 0) & "��-��" & teamNameOfNextSection(1, 1) & " �ŁA" & vbCrLf & _
                                       "�@" & Sheets(seasonName & "_�e��L�^").Cells(teamID + 2, "BA").Value & "�̎��͗D�������Łc"
                End If
                
            Next teamID
            
            ' ����V����
            For teamID = 0 To 4
                
                If Left(vStatus(1, teamID), 1) = "��" And Left(vStatus(2, teamID), 1) = "��" And Left(vStatus(4, teamID), 1) <> "��" And Left(vStatus(5, teamID), 1) = "��" Then
                    rankNotification = rankNotification & vbCrLf & _
                                       "�E���� " & teamNameOfNextSection(0, 0) & "��-��" & teamNameOfNextSection(0, 1) & " & " & teamNameOfNextSection(1, 0) & "��-��" & teamNameOfNextSection(1, 1) & " �ŁA" & vbCrLf & _
                                       "�@" & Sheets(seasonName & "_�e��L�^").Cells(teamID + 2, "BA").Value & "�̎��͗D���������I"
                End If
                
            Next teamID
            
            ' �}�W�b�N�_��
            For teamID = 0 To 4
                
                If Left(vStatus(1, teamID), 1) <> "M" And Left(vStatus(2, teamID), 1) <> "M" And Left(vStatus(4, teamID), 1) = "M" And Left(vStatus(5, teamID), 1) <> "M" Then
                    rankNotification = rankNotification & vbCrLf & _
                                       "�E���� " & teamNameOfNextSection(0, 0) & "��-��" & teamNameOfNextSection(0, 1) & " & " & teamNameOfNextSection(1, 0) & "��-��" & teamNameOfNextSection(1, 1) & " �ŁA" & vbCrLf & _
                                       "�@" & Sheets(seasonName & "_�e��L�^").Cells(teamID + 2, "BA").Value & "�̗D���}�W�b�N���_���I"
                End If
                
            Next teamID
            
        End If
        
MPB_NEWS_PATTERN7_END_POINT:

        ' �@��-���A��-��
        If vStatus(5, 5) = "true" Then
            ' �D��
            For teamID = 0 To 4
                
                If Left(vStatus(1, teamID), 1) <> "�D" And Left(vStatus(3, teamID), 1) <> "�D" And Left(vStatus(4, teamID), 1) <> "�D" And Left(vStatus(5, teamID), 1) = "�D" Then
                    rankNotification = rankNotification & vbCrLf & _
                                       "�E���� " & teamNameOfNextSection(0, 0) & "��-��" & teamNameOfNextSection(0, 1) & " & " & teamNameOfNextSection(1, 0) & "��-��" & teamNameOfNextSection(1, 1) & " �ŁA" & vbCrLf & _
                                       "�@" & Sheets(seasonName & "_�e��L�^").Cells(teamID + 2, "BA").Value & "�̃��[�O�D�����m��I"
                    GoTo MPB_NEWS_PATTERN8_END_POINT
                End If
                
            Next teamID
            
            ' �}�W�b�N����
            For teamID = 0 To 4
                
                If Left(vStatus(1, teamID), 1) = "M" And Left(vStatus(3, teamID), 1) = "M" And Left(vStatus(4, teamID), 1) = "M" And Left(vStatus(5, teamID), 1) <> "M" Then
                    rankNotification = rankNotification & vbCrLf & _
                                       "�E���� " & teamNameOfNextSection(0, 0) & "��-��" & teamNameOfNextSection(0, 1) & " & " & teamNameOfNextSection(1, 0) & "��-��" & teamNameOfNextSection(1, 1) & " �ŁA" & vbCrLf & _
                                       "�@" & Sheets(seasonName & "_�e��L�^").Cells(teamID + 2, "BA").Value & "�̗D���}�W�b�N�����Łc"
                End If
                
            Next teamID
            
            ' ����V����
            For teamID = 0 To 4
                
                If Left(vStatus(1, teamID), 1) <> "��" And Left(vStatus(3, teamID), 1) <> "��" And Left(vStatus(4, teamID), 1) <> "��" And Left(vStatus(5, teamID), 1) = "��" Then
                    rankNotification = rankNotification & vbCrLf & _
                                       "�E���� " & teamNameOfNextSection(0, 0) & "��-��" & teamNameOfNextSection(0, 1) & " & " & teamNameOfNextSection(1, 0) & "��-��" & teamNameOfNextSection(1, 1) & " �ŁA" & vbCrLf & _
                                       "�@" & Sheets(seasonName & "_�e��L�^").Cells(teamID + 2, "BA").Value & "�̎��͗D�������Łc"
                End If
                
            Next teamID
            
            ' ����V����
            For teamID = 0 To 4
                
                If Left(vStatus(1, teamID), 1) = "��" And Left(vStatus(3, teamID), 1) = "��" And Left(vStatus(4, teamID), 1) = "��" And Left(vStatus(5, teamID), 1) <> "��" Then
                    rankNotification = rankNotification & vbCrLf & _
                                       "�E���� " & teamNameOfNextSection(0, 0) & "��-��" & teamNameOfNextSection(0, 1) & " & " & teamNameOfNextSection(1, 0) & "��-��" & teamNameOfNextSection(1, 1) & " �ŁA" & vbCrLf & _
                                       "�@" & Sheets(seasonName & "_�e��L�^").Cells(teamID + 2, "BA").Value & "�̎��͗D���������I"
                End If
                
            Next teamID
            
            ' �}�W�b�N�_��
            For teamID = 0 To 4
                
                If Left(vStatus(1, teamID), 1) <> "M" And Left(vStatus(3, teamID), 1) <> "M" And Left(vStatus(4, teamID), 1) <> "M" And Left(vStatus(5, teamID), 1) = "M" Then
                    rankNotification = rankNotification & vbCrLf & _
                                       "�E���� " & teamNameOfNextSection(0, 0) & "��-��" & teamNameOfNextSection(0, 1) & " & " & teamNameOfNextSection(1, 0) & "��-��" & teamNameOfNextSection(1, 1) & " �ŁA" & vbCrLf & _
                                       "�@" & Sheets(seasonName & "_�e��L�^").Cells(teamID + 2, "BA").Value & "�̗D���}�W�b�N���_���I"
                End If
                
            Next teamID
            
        End If
        
MPB_NEWS_PATTERN8_END_POINT:
        
    End If
    
MPB_NEWS_CHECK_END_POINT:
    
    ' TS/OB�g�U�蒼���̂��m�点
    If numberOfSection = 10 Or numberOfSection = 20 Then
        tsobChangeNotification = tsobChangeNotification & vbCrLf & _
                                 "��" & numberOfSection & "�߂��I�������̂ŁATS/OB�g�̐U�蒼�����s���܂��B" & vbCrLf & _
                                 "- - - - - - - - - -" & vbCrLf & _
                                 "1��:" & Left(Sheets(seasonName & "_�e��L�^").Cells(2, "B").Value, 1) & "�@���@3.5" & vbCrLf & _
                                 "2��:" & Left(Sheets(seasonName & "_�e��L�^").Cells(3, "B").Value, 1) & "�@���@4.0" & vbCrLf & _
                                 "3��:" & Left(Sheets(seasonName & "_�e��L�^").Cells(4, "B").Value, 1) & "�@���@4.5" & vbCrLf & _
                                 "4��:" & Left(Sheets(seasonName & "_�e��L�^").Cells(5, "B").Value, 1) & "�@���@5.0" & vbCrLf & _
                                 "5��:" & Left(Sheets(seasonName & "_�e��L�^").Cells(6, "B").Value, 1) & "�@���@5.5" & vbCrLf & _
                                 "�������������͂��̒l�ǂ���̕ύX�ɂȂ�Ȃ��ꍇ������܂��B���m�ȏ������҂����������B" & vbCrLf & _
                                 "- - - - - - - - - -" & vbCrLf & _
                                 "�ȏ�"
    End If
    
    Dim BASIC_CLUSTER_RATIO As Single
    BASIC_CLUSTER_RATIO = 0
    Dim SMALL_CLUSTER_RATIO, BIG_CLUSTER_RATIO As Integer
    SMALL_CLUSTER_RATIO = 70
    BIG_CLUSTER_RATIO = 30
    If SMALL_CLUSTER_RATIO + BIG_CLUSTER_RATIO <> 100 Then
        MsgBox "CLUSTER_RATIO�̐ݒ肪�s���ł��B"
        End
    End If
    Dim PCR_POSITIVE_RATIO_IN_SMALL_CLUSTER, PCR_POSITIVE_RATIO_IN_BIG_CLUSTER As Single
    PCR_POSITIVE_RATIO_IN_SMALL_CLUSTER = 0.1
    PCR_POSITIVE_RATIO_IN_BIG_CLUSTER = 0.3
    
    Dim BASIC_ACCIDENT_RATIO As Single
    BASIC_ACCIDENT_RATIO = 0.007
    Dim TWO_GAMES_ACCIDENT_RATIO, FIVE_GAMES_ACCIDENT_RATIO, EIGHT_GAMES_ACCIDENT_RATIO, ALL_GAMES_ACCIDENT_RATIO As Integer
    TWO_GAMES_ACCIDENT_RATIO = 70
    FIVE_GAMES_ACCIDENT_RATIO = 15
    EIGHT_GAMES_ACCIDENT_RATIO = 10
    ALL_GAMES_ACCIDENT_RATIO = 5
    If TWO_GAMES_ACCIDENT_RATIO + FIVE_GAMES_ACCIDENT_RATIO + EIGHT_GAMES_ACCIDENT_RATIO + ALL_GAMES_ACCIDENT_RATIO <> 100 Then
        MsgBox "ACCIDENT_RATIO�̐ݒ肪�s���ł��B"
        End
    End If
    Dim ACCIDENT_PERIOD_SHORT_RATIO, ACCIDENT_PERIOD_NORMAL_RATIO, ACCIDENT_PERIOD_LONG_RATIO As Integer
    ACCIDENT_PERIOD_SHORT_RATIO = 30
    ACCIDENT_PERIOD_NORMAL_RATIO = 40
    ACCIDENT_PERIOD_LONG_RATIO = 30
    If ACCIDENT_PERIOD_SHORT_RATIO + ACCIDENT_PERIOD_NORMAL_RATIO + ACCIDENT_PERIOD_LONG_RATIO <> 100 Then
        MsgBox "ACCIDENT_PERIOD_RATIO�̐ݒ肪�s���ł��B"
        End
    End If
    
    Dim ACCIDENT_COEFFICIENT_S, ACCIDENT_COEFFICIENT_A, ACCIDENT_COEFFICIENT_B, ACCIDENT_COEFFICIENT_C, _
        ACCIDENT_COEFFICIENT_D, ACCIDENT_COEFFICIENT_E, ACCIDENT_COEFFICIENT_F, ACCIDENT_COEFFICIENT_G, ACCIDENT_COEFFICIENT_n As Single
    ACCIDENT_COEFFICIENT_S = 0.01
    ACCIDENT_COEFFICIENT_A = 0.3
    ACCIDENT_COEFFICIENT_B = 0.5
    ACCIDENT_COEFFICIENT_C = 0.8
    ACCIDENT_COEFFICIENT_D = 1#
    ACCIDENT_COEFFICIENT_E = 1.2
    ACCIDENT_COEFFICIENT_F = 2#
    ACCIDENT_COEFFICIENT_G = 4#
    ACCIDENT_COEFFICIENT_n = 0#
    
    Sheets(seasonName & "_����f�[�^").Unprotect
    Sheets(seasonName & "_���f�[�^").Unprotect
    
    For teamID = 0 To 4
        ' �ϐ��̏�����
        teamName = ""
        numberOfGamesPlayedBeforeThisSection = 0
        numberOfGamesPlayedAfterThisSection = 0
        pcrPositiveRate = 0#
        accidentBonusValue = 1#
        pcrResultText = ""
        pcrResultMessage = ""
        
        Sheets(seasonName & "_�X�P�W���[��").Activate
        
        ' �����i�s�󋵂̊m�F
        numberOfGamesPlayedAfterThisSection = ActiveSheet.Cells(2 + numberOfSection, 84 + teamID)
        
        If numberOfSection = 0 Then
            numberOfGamesPlayedBeforeThisSection = -100
        Else
            numberOfGamesPlayedBeforeThisSection = ActiveSheet.Cells(1 + numberOfSection, 84 + teamID)
        End If
        
        Select Case ActiveSheet.Cells(1, 54 + teamID)
            Case Is = "G"
                teamName = "�W���C�A���c"
                accidentBonusValue = 1#
            Case Is = "L"
                teamName = "���C�I���Y"
                accidentBonusValue = 1#
            Case Is = "E"
                teamName = "�C�[�O���X"
                accidentBonusValue = 1#
            Case Is = "T"
                teamName = "�^�C�K�[�X"
                accidentBonusValue = 1#
            Case Is = "M"
                teamName = "�}���[���Y"
                accidentBonusValue = 1#
            Case Else
                MsgBox "��O���������܂����i1001�j"
                End
        End Select
        
        ' �X�y����J�n
        If debugModeFlg Then
            MsgBox ActiveSheet.Cells(1, 54 + teamID) & "�F����J�n"
        End If
        
        ' �N���X�^�[����
        dice = Rnd()
        If dice < BASIC_CLUSTER_RATIO Then
            dice = Rnd() * 100
            Select Case dice
                Case Is < SMALL_CLUSTER_RATIO
                    pcrPositiveRate = PCR_POSITIVE_RATIO_IN_SMALL_CLUSTER
                Case Is < SMALL_CLUSTER_RATIO + BIG_CLUSTER_RATIO
                    pcrPositiveRate = PCR_POSITIVE_RATIO_IN_BIG_CLUSTER
                Case Else
                    MsgBox "��O���������܂����i1002�j"
                    End
            End Select
            pcrResultMessage = teamName & "�ɂăN���X�^�[���������܂����B����ɂ��o�^�����I��͎��̒ʂ�ł��B�F"
            pcrResultText = vbCrLf & _
                            "��" & teamName & "������X�N���[�j���O�����ŋ��c�֌W�҂��܂ރN���X�^�[�����o�B����ɂ�莟�̑I�肪�o�^�����ƂȂ�܂����B"
        Else
            pcrPositiveRate = 0#
        End If
        
        ' ���藣�E����
        Sheets(seasonName & "_����f�[�^").Activate
        pcrResultText = pcrResultText & vbCrLf & _
                        "�i����j"
        For playerID = 4 To 50
            ' �ϐ��̏�����
            rowIdx = teamID * 50 + playerID
            playerName = ActiveSheet.Cells(rowIdx, "B").Value
            playerNameRegistered = ActiveSheet.Cells(rowIdx, "D").Value
            accidentRate = BASIC_ACCIDENT_RATIO * accidentBonusValue
            accidentPeriod = 0
            accidentPeriodString = ""
            accidentOverview = ""
            accidentText = ""
            accidentMessage = ""
            
            If playerName = "" Then
                GoTo STATUS_n_POINT_1
            End If
            If ActiveSheet.Cells(rowIdx, 282 + numberOfGamesPlayedAfterThisSection).Value <> "" Or numberOfGamesPlayedAfterThisSection = numberOfGamesPlayedBeforeThisSection Then
                GoTo ACCIDENT_PERIOD_ZERO_POINT_1
            End If
            ' �X�y����
            ' ��b�X�y��
            Select Case ActiveSheet.Cells(rowIdx, "E").Value
                Case Is = "S"
                    accidentRate = accidentRate * ACCIDENT_COEFFICIENT_S
                Case Is = "A"
                    accidentRate = accidentRate * ACCIDENT_COEFFICIENT_A
                Case Is = "B"
                    accidentRate = accidentRate * ACCIDENT_COEFFICIENT_B
                Case Is = "C"
                    accidentRate = accidentRate * ACCIDENT_COEFFICIENT_C
                Case Is = "D"
                    accidentRate = accidentRate * ACCIDENT_COEFFICIENT_D
                Case Is = "E"
                    accidentRate = accidentRate * ACCIDENT_COEFFICIENT_E
                Case Is = "F"
                    accidentRate = accidentRate * ACCIDENT_COEFFICIENT_F
                Case Is = "G"
                    accidentRate = accidentRate * ACCIDENT_COEFFICIENT_G
                Case Is = "n"
                    accidentRate = accidentRate * ACCIDENT_COEFFICIENT_n
                    GoTo STATUS_n_POINT_1
                Case Else
                    MsgBox "��O���������܂����i1101�j"
                    End
            End Select
                    
            ' �~����J>=120�ŃX�y����10�{���鏈���i����̂݁j
            columnIdxOfStamina = 161 + numberOfGamesPlayedAfterThisSection * 5
            If ActiveSheet.Cells(rowIdx, columnIdxOfStamina).Value - ActiveSheet.Cells(rowIdx, columnIdxOfStamina - 2).Value >= 120 Then
                accidentRate = accidentRate * 10
            End If
            
            ' �����i�s�ɔ����ăX�y�����㏸���鏈��
            accidentRate = accidentRate * (0.885 + (numberOfGamesPlayedAfterThisSection * 0.01))
              
            ' �X�y�d���̌���E�X�y�d���̃����_���v�f
            dice = Rnd()
            If dice < accidentRate Then
                dice = Rnd() * 100
                Select Case dice
                    Case Is < TWO_GAMES_ACCIDENT_RATIO
                        accidentPeriod = 2
                        accidentPeriodString = "2"
                    Case Is < TWO_GAMES_ACCIDENT_RATIO + FIVE_GAMES_ACCIDENT_RATIO
                        accidentPeriod = 5
                        accidentPeriodString = "5"
                    Case Is < TWO_GAMES_ACCIDENT_RATIO + FIVE_GAMES_ACCIDENT_RATIO + EIGHT_GAMES_ACCIDENT_RATIO
                        accidentPeriod = 8
                        accidentPeriodString = "8"
                    Case Is < TWO_GAMES_ACCIDENT_RATIO + FIVE_GAMES_ACCIDENT_RATIO + EIGHT_GAMES_ACCIDENT_RATIO + ALL_GAMES_ACCIDENT_RATIO
                        accidentPeriod = 24
                        accidentPeriodString = "-"
                        GoTo ACCIDENT_PERIOD_ALL_POINT_1
                End Select
                
                dice = Rnd() * 100
                Select Case dice
                    Case Is < ACCIDENT_PERIOD_SHORT_RATIO
                        accidentPeriod = accidentPeriod - 1
                    Case Is < ACCIDENT_PERIOD_SHORT_RATIO + ACCIDENT_PERIOD_NORMAL_RATIO
                        accidentPeriod = accidentPeriod
                    Case Is < ACCIDENT_PERIOD_SHORT_RATIO + ACCIDENT_PERIOD_NORMAL_RATIO + ACCIDENT_PERIOD_LONG_RATIO
                        accidentPeriod = accidentPeriod + 1
                    Case Else
                        MsgBox "��O���������܂����i1102�j"
                        End
                End Select
            Else
                GoTo ACCIDENT_PERIOD_ZERO_POINT_1
            End If
            
ACCIDENT_PERIOD_ALL_POINT_1:
    
            ' ��̓I�ȃX�y�̌���
            dice = Rnd()
            accidentOverview = playerNameRegistered & ":" & Sheets("�A�N�V�f���g").Cells(Int(5 * dice) + 2, "A").Value & "(" & accidentPeriodString & ")"
            Select Case accidentPeriodString
                Case Is = "2"
                    accidentMessage = playerName & " �I��F" & vbCrLf & _
                                      Sheets("�A�N�V�f���g").Cells(Int(5 * dice) + 2, "B").Value & vbCrLf & _
                                      Sheets("�A�N�V�f���g").Cells(12, "B").Value
                    allAccidentResultNotification = allAccidentResultNotification & vbCrLf & _
                                                    "��" & teamName & "��" & playerNameRegistered & "�I�肪" & Sheets("�A�N�V�f���g").Cells(Int(5 * dice) + 2, "B").Value & Sheets("�A�N�V�f���g").Cells(12, "B").Value
                Case Is = "5"
                    accidentMessage = playerName & " �I��F" & vbCrLf & _
                                      Sheets("�A�N�V�f���g").Cells(Int(5 * dice) + 2, "C").Value & vbCrLf & _
                                      Sheets("�A�N�V�f���g").Cells(12, "C").Value
                    allAccidentResultNotification = allAccidentResultNotification & vbCrLf & _
                                                    "��" & teamName & "��" & playerNameRegistered & "�I�肪" & Sheets("�A�N�V�f���g").Cells(Int(5 * dice) + 2, "C").Value & Sheets("�A�N�V�f���g").Cells(12, "C").Value
                Case Is = "8"
                    accidentMessage = playerName & " �I��F" & vbCrLf & _
                                      Sheets("�A�N�V�f���g").Cells(Int(5 * dice) + 2, "D").Value & vbCrLf & _
                                      Sheets("�A�N�V�f���g").Cells(12, "D").Value
                    allAccidentResultNotification = allAccidentResultNotification & vbCrLf & _
                                                    "��" & teamName & "��" & playerNameRegistered & "�I�肪" & Sheets("�A�N�V�f���g").Cells(Int(5 * dice) + 2, "D").Value & Sheets("�A�N�V�f���g").Cells(12, "D").Value
                Case Is = "-"
                    accidentMessage = playerName & " �I��F" & vbCrLf & _
                                      Sheets("�A�N�V�f���g").Cells(Int(5 * dice) + 2, "E").Value & vbCrLf & _
                                      Sheets("�A�N�V�f���g").Cells(12, "E").Value
                    allAccidentResultNotification = allAccidentResultNotification & vbCrLf & _
                                                    "��" & teamName & "��" & playerNameRegistered & "�I�肪" & Sheets("�A�N�V�f���g").Cells(Int(5 * dice) + 2, "E").Value & Sheets("�A�N�V�f���g").Cells(12, "E").Value
                Case Else
                    MsgBox "��O���������܂����i1103�j"
                    End
            End Select
            
            ' ��������
            If debugModeFlg Then
                MsgBox accidentMessage
            End If
            
            For columnIdx = 0 To accidentPeriod - 1
                If 282 + numberOfGamesPlayedAfterThisSection + columnIdx <= 305 Then
                    ActiveSheet.Cells(rowIdx, 282 + numberOfGamesPlayedAfterThisSection + columnIdx).Value = accidentOverview
                End If
            Next columnIdx
            
ACCIDENT_PERIOD_ZERO_POINT_1:
    
            ' ���ᔻ��
            dice = Rnd()
            If dice < pcrPositiveRate Then
                pcrResultMessage = pcrResultMessage & vbCrLf & _
                                   "�E" & playerName
                pcrResultText = pcrResultText & playerNameRegistered & " "
                If 282 + numberOfGamesPlayedAfterThisSection <= 305 Then
                    ActiveSheet.Cells(rowIdx, 282 + numberOfGamesPlayedAfterThisSection).Value = playerNameRegistered & ":����"
                End If
            End If
            
            ' ���A����
            If numberOfSection > 0 And ActiveSheet.Cells(rowIdx, 282 + numberOfGamesPlayedBeforeThisSection).Value <> "" And ActiveSheet.Cells(rowIdx, 282 + numberOfGamesPlayedAfterThisSection).Value = "" Then
            
                If debugModeFlg Then
                    MsgBox playerName & " �I��F" & vbCrLf & _
                           "���߂���̐�񕜋A����]�w�ɂ���Ė�������܂����B"
                End If
                
                allAccidentResultNotification = allAccidentResultNotification & vbCrLf & _
                                                "��" & teamName & "�����E����" & playerNameRegistered & "�I��Ɋւ��A���߂���̐�񕜋A����������܂����B"
            End If
            
STATUS_n_POINT_1:
        
        Next playerID
        
        ' ��藣�E����
        Sheets(seasonName & "_���f�[�^").Activate
        pcrResultText = pcrResultText & vbCrLf & _
                        "�i���j"
        For playerID = 4 To 50
            ' �ϐ��̏�����
            rowIdx = teamID * 50 + playerID
            playerName = ActiveSheet.Cells(rowIdx, "B").Value
            playerNameRegistered = ActiveSheet.Cells(rowIdx, "D").Value
            accidentRate = BASIC_ACCIDENT_RATIO
            accidentPeriod = 0
            accidentPeriodString = ""
            accidentOverview = ""
            accidentText = ""
            accidentMessage = ""
            
            If playerName = "" Then
                GoTo STATUS_n_POINT_2
            End If
            If ActiveSheet.Cells(rowIdx, 236 + numberOfGamesPlayedAfterThisSection).Value <> "" Or numberOfGamesPlayedAfterThisSection = numberOfGamesPlayedBeforeThisSection Then
                GoTo ACCIDENT_PERIOD_ZERO_POINT_2
            End If
            ' �X�y����
            ' ��b�X�y��
            Select Case ActiveSheet.Cells(rowIdx, "E").Value
                Case Is = "S"
                    accidentRate = accidentRate * ACCIDENT_COEFFICIENT_S
                Case Is = "A"
                    accidentRate = accidentRate * ACCIDENT_COEFFICIENT_A
                Case Is = "B"
                    accidentRate = accidentRate * ACCIDENT_COEFFICIENT_B
                Case Is = "C"
                    accidentRate = accidentRate * ACCIDENT_COEFFICIENT_C
                Case Is = "D"
                    accidentRate = accidentRate * ACCIDENT_COEFFICIENT_D
                Case Is = "E"
                    accidentRate = accidentRate * ACCIDENT_COEFFICIENT_E
                Case Is = "F"
                    accidentRate = accidentRate * ACCIDENT_COEFFICIENT_F
                Case Is = "G"
                    accidentRate = accidentRate * ACCIDENT_COEFFICIENT_G
                Case Is = "n"
                    accidentRate = accidentRate * ACCIDENT_COEFFICIENT_n
                    GoTo STATUS_n_POINT_2
                Case Else
                    MsgBox "��O���������܂����i1201�j"
                    End
            End Select
                    
            ' �����i�s�ɔ����ăX�y�����㏸���鏈��
            accidentRate = accidentRate * (0.885 + (numberOfGamesPlayedAfterThisSection * 0.01))
            
            ' �X�y�d���̌���E�X�y�d���̃����_���v�f
            dice = Rnd()
            If dice < accidentRate Then
                dice = Rnd() * 100
                Select Case dice
                    Case Is < TWO_GAMES_ACCIDENT_RATIO
                        accidentPeriod = 2
                        accidentPeriodString = "2"
                    Case Is < TWO_GAMES_ACCIDENT_RATIO + FIVE_GAMES_ACCIDENT_RATIO
                        accidentPeriod = 5
                        accidentPeriodString = "5"
                    Case Is < TWO_GAMES_ACCIDENT_RATIO + FIVE_GAMES_ACCIDENT_RATIO + EIGHT_GAMES_ACCIDENT_RATIO
                        accidentPeriod = 8
                        accidentPeriodString = "8"
                    Case Is < TWO_GAMES_ACCIDENT_RATIO + FIVE_GAMES_ACCIDENT_RATIO + EIGHT_GAMES_ACCIDENT_RATIO + ALL_GAMES_ACCIDENT_RATIO
                        accidentPeriod = 24
                        accidentPeriodString = "-"
                        GoTo ACCIDENT_PERIOD_ALL_POINT_2
                End Select
                
                dice = Rnd() * 100
                Select Case dice
                    Case Is < ACCIDENT_PERIOD_SHORT_RATIO
                        accidentPeriod = accidentPeriod - 1
                    Case Is < ACCIDENT_PERIOD_SHORT_RATIO + ACCIDENT_PERIOD_NORMAL_RATIO
                        accidentPeriod = accidentPeriod
                    Case Is < ACCIDENT_PERIOD_SHORT_RATIO + ACCIDENT_PERIOD_NORMAL_RATIO + ACCIDENT_PERIOD_LONG_RATIO
                        accidentPeriod = accidentPeriod + 1
                    Case Else
                        MsgBox "��O���������܂����i1202�j"
                        End
                End Select
            Else
                GoTo ACCIDENT_PERIOD_ZERO_POINT_2
            End If
            
ACCIDENT_PERIOD_ALL_POINT_2:
    
            ' ��̓I�ȃX�y�̌���
            dice = Rnd()
            accidentOverview = playerNameRegistered & ":" & Sheets("�A�N�V�f���g").Cells(Int(5 * dice) + 2, "A").Value & "(" & accidentPeriodString & ")"
            Select Case accidentPeriodString
                Case Is = "2"
                    accidentMessage = playerName & " �I��F" & vbCrLf & _
                                      Sheets("�A�N�V�f���g").Cells(Int(4 * dice) + 8, "B").Value & vbCrLf & _
                                      Sheets("�A�N�V�f���g").Cells(12, "B").Value
                    allAccidentResultNotification = allAccidentResultNotification & vbCrLf & _
                                                    "��" & teamName & "��" & playerNameRegistered & "�I�肪" & Sheets("�A�N�V�f���g").Cells(Int(5 * dice) + 2, "B").Value & Sheets("�A�N�V�f���g").Cells(12, "B").Value
                Case Is = "5"
                    accidentMessage = playerName & " �I��F" & vbCrLf & _
                                      Sheets("�A�N�V�f���g").Cells(Int(4 * dice) + 8, "C").Value & vbCrLf & _
                                      Sheets("�A�N�V�f���g").Cells(12, "C").Value
                    allAccidentResultNotification = allAccidentResultNotification & vbCrLf & _
                                                    "��" & teamName & "��" & playerNameRegistered & "�I�肪" & Sheets("�A�N�V�f���g").Cells(Int(5 * dice) + 2, "C").Value & Sheets("�A�N�V�f���g").Cells(12, "C").Value
                Case Is = "8"
                    accidentMessage = playerName & " �I��F" & vbCrLf & _
                                      Sheets("�A�N�V�f���g").Cells(Int(4 * dice) + 8, "D").Value & vbCrLf & _
                                      Sheets("�A�N�V�f���g").Cells(12, "D").Value
                    allAccidentResultNotification = allAccidentResultNotification & vbCrLf & _
                                                    "��" & teamName & "��" & playerNameRegistered & "�I�肪" & Sheets("�A�N�V�f���g").Cells(Int(5 * dice) + 2, "D").Value & Sheets("�A�N�V�f���g").Cells(12, "D").Value
                Case Is = "-"
                    accidentMessage = playerName & " �I��F" & vbCrLf & _
                                      Sheets("�A�N�V�f���g").Cells(Int(4 * dice) + 8, "E").Value & vbCrLf & _
                                      Sheets("�A�N�V�f���g").Cells(12, "E").Value
                    allAccidentResultNotification = allAccidentResultNotification & vbCrLf & _
                                                    "��" & teamName & "��" & playerNameRegistered & "�I�肪" & Sheets("�A�N�V�f���g").Cells(Int(5 * dice) + 2, "E").Value & Sheets("�A�N�V�f���g").Cells(12, "E").Value
                Case Else
                    MsgBox "��O���������܂����i1203�j"
                    End
            End Select
            
            ' ��������
            If debugModeFlg Then
                MsgBox accidentMessage
            End If
            
            For columnIdx = 0 To accidentPeriod - 1
                If 236 + numberOfGamesPlayedAfterThisSection + columnIdx <= 259 Then
                    ActiveSheet.Cells(rowIdx, 236 + numberOfGamesPlayedAfterThisSection + columnIdx).Value = accidentOverview
                End If
            Next columnIdx
            
ACCIDENT_PERIOD_ZERO_POINT_2:
    
            ' ���ᔻ��
            dice = Rnd()
            If dice < pcrPositiveRate Then
                pcrResultMessage = pcrResultMessage & vbCrLf & _
                                   "�E" & playerName
                pcrResultText = pcrResultText & playerNameRegistered & " "
                If 236 + numberOfGamesPlayedAfterThisSection <= 259 Then
                    ActiveSheet.Cells(rowIdx, 236 + numberOfGamesPlayedAfterThisSection).Value = playerNameRegistered & ":����"
                End If
            End If
            
            ' ���A����
            If numberOfSection > 0 And ActiveSheet.Cells(rowIdx, 236 + numberOfGamesPlayedBeforeThisSection).Value <> "" And ActiveSheet.Cells(rowIdx, 236 + numberOfGamesPlayedAfterThisSection).Value = "" Then
                
                If debugModeFlg Then
                    MsgBox playerName & " �I��F" & vbCrLf & _
                           "���߂���̐�񕜋A����]�w�ɂ���Ė�������܂����B"
                End If
                
                allAccidentResultNotification = allAccidentResultNotification & vbCrLf & _
                                                "��" & teamName & "�����E����" & playerNameRegistered & "�I��Ɋւ��A���߂���̐�񕜋A����������܂����B"
            End If
            
STATUS_n_POINT_2:
    
        Next playerID
            
        ' �N���X�^�[����̌��ʂ��o��
        If pcrPositiveRate > 0 Then
            ' MsgBox pcrResultMessage
            allAccidentResultNotification = allAccidentResultNotification & pcrResultText
        End If
        
        ' �I���̏o��
        Sheets(seasonName & "_�X�P�W���[��").Activate
        If debugModeFlg Then
            If numberOfGamesPlayedAfterThisSection > numberOfGamesPlayedBeforeThisSection Then
                MsgBox ActiveSheet.Cells(1, 54 + teamID) & "�F����I��"
            Else
                MsgBox ActiveSheet.Cells(1, 54 + teamID) & "�F����I��(����̂�)"
            End If
        End If
    Next teamID
    
    Sheets(seasonName & "_����f�[�^").Protect AllowFormattingColumns:=True, AllowFormattingRows:=True
    Sheets(seasonName & "_���f�[�^").Protect AllowFormattingColumns:=True, AllowFormattingRows:=True
    Sheets(seasonName & "_�X�P�W���[��").Activate
    
    If Not debugModeFlg Then
        Call �o�b�N�A�b�v
        
        Open "C:\Users\TaiNo\�}�C�h���C�u\�������A�^�C��_���M�ҋ@\accident.txt" For Output As #1
            Print #1, allAccidentResultNotification;
        Close #1
        
        If rankNotification <> "�yMPB�j���[�X�z" Then
            Open "C:\Users\TaiNo\�}�C�h���C�u\�������A�^�C��_���M�ҋ@\news.txt" For Output As #3
                Print #3, rankNotification;
            Close #3
        End If
        
        If tsobChangeNotification <> "�yTS/OB�g�U�蒼���̂��m�点�z" Then
            Open "C:\Users\TaiNo\�}�C�h���C�u\�������A�^�C��_���M�ҋ@\tsob.txt" For Output As #4
                Print #4, tsobChangeNotification;
            Close #4
        End If
        
        Call �摜�ۑ�
        
        Application.ScreenUpdating = True
        ActiveWorkbook.Close SaveChanges:=True
    Else
        MsgBox allAccidentResultNotification
        MsgBox rankNotification
        MsgBox tsobChangeNotification

        MsgBox "�������I�����܂�"
    End If
    
End Sub


