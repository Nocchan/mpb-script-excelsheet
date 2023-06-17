Attribute VB_Name = "mpb_vbascript_sortPlayerSheet"
Sub ���ёւ�()

    ' �G���[�`�F�b�N
    If ActiveSheet.Name <> ActiveSheet.Cells(1, "H").Value & "_����f�[�^" And ActiveSheet.Name <> ActiveSheet.Cells(1, "H").Value & "_���f�[�^" Then
        If ActiveSheet.Name <> "�L�^��_" & ActiveSheet.Cells(2, "A").Value Then
            MsgBox "�V�[�g�����s���ł��B"
            End
        End If
    End If

    Application.ScreenUpdating = False
    ActiveSheet.Unprotect

    ' ����/���f�[�^�̏ꍇ
    If ActiveSheet.Name = ActiveSheet.Cells(1, "H").Value & "_����f�[�^" Or ActiveSheet.Name = ActiveSheet.Cells(1, "H").Value & "_���f�[�^" Then
    
        With ActiveSheet.Sort
            .SortFields.Clear
            .SortFields.Add2 Key:=Range("A4:A50"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            .SortFields.Add2 Key:=Range("B4:B50"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            .SetRange Range("A4:ZZ50")
            .Header = xlGuess
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
            
            .SortFields.Add2 Key:=Range("A54:A100"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            .SortFields.Add2 Key:=Range("B54:B100"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            .SetRange Range("A54:ZZ100")
            .Header = xlGuess
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
            
            .SortFields.Add2 Key:=Range("A104:A150"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            .SortFields.Add2 Key:=Range("B104:B150"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            .SetRange Range("A104:ZZ150")
            .Header = xlGuess
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
            
            .SortFields.Add2 Key:=Range("A154:A200"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            .SortFields.Add2 Key:=Range("B154:B200"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            .SetRange Range("A154:ZZ200")
            .Header = xlGuess
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
            
            .SortFields.Add2 Key:=Range("A204:A250"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            .SortFields.Add2 Key:=Range("B204:B250"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            .SetRange Range("A204:ZZ250")
            .Header = xlGuess
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
        
        ActiveSheet.Range("$A$1:$A$250").AutoFilter Field:=1, Criteria1:="<>", VisibleDropDown:=False

    End If
    
    ' �L�^���̏ꍇ
    If ActiveSheet.Name = "�L�^��_" & ActiveSheet.Cells(2, "A").Value Then

        Dim rowIndex As Integer
        
        ActiveSheet.Rows("5:103").Hidden = False
        ActiveSheet.Rows("105:203").Hidden = False
        
        With ActiveSheet.Sort
            .SortFields.Clear
    
            .SortFields.Add2 Key:=Range("B5:B103"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            .SortFields.Add2 Key:=Range("C5:C103"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            .SetRange Range("B4:BU103") ' �V�[�Y���X�V���C��1/2
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
    
            .SortFields.Add2 Key:=Range("B105:B203"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            .SortFields.Add2 Key:=Range("C105:C203"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            .SetRange Range("B104:BU203") ' �V�[�Y���X�V���C��2/2
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
        
        For rowIndex = 5 To 102
            If ActiveSheet.Cells(rowIndex, "B").Value = "" Then
                ActiveSheet.Rows(WorksheetFunction.Min(rowIndex, 102) & ":102").Hidden = True
                Exit For
            End If
        Next rowIndex
        
        For rowIndex = 105 To 202
            If ActiveSheet.Cells(rowIndex, "B").Value = "" Then
                ActiveSheet.Rows(WorksheetFunction.Min(rowIndex, 202) & ":202").Hidden = True
                Exit For
            End If
        Next rowIndex

    End If
    
    Call �o�b�N�A�b�v
    
    ActiveSheet.Protect AllowFormattingColumns:=True, AllowFormattingRows:=True
    Application.ScreenUpdating = True
    
End Sub


