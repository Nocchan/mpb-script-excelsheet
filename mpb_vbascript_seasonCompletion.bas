Attribute VB_Name = "mpb_vbascript_seasonCompletion"
Sub �X�y�󋵃��Z�b�g()

    ' �G���[�`�F�b�N
    If ActiveSheet.Name <> ActiveSheet.Cells(1, "A").Value & "_�X�P�W���[��" Then
        MsgBox "�V�[�g���܂���A1�Z���̃V�[�Y���w�肪�s���ł��B"
        End
    End If

    Application.ScreenUpdating = False
    
    seasonName = ActiveSheet.Cells(1, "A")
    
    With Sheets(seasonName & "_����f�[�^")
        .Unprotect
        .Range("JV4:KS50").ClearContents
        .Range("JV54:KS100").ClearContents
        .Range("JV104:KS150").ClearContents
        .Range("JV154:KS200").ClearContents
        .Range("JV204:KS250").ClearContents
        .Protect AllowFormattingColumns:=True, AllowFormattingRows:=True
    End With
    
    With Sheets(seasonName & "_���f�[�^")
        .Unprotect
        .Range("IB4:IY50").ClearContents
        .Range("IB54:IY100").ClearContents
        .Range("IB104:IY150").ClearContents
        .Range("IB154:IY200").ClearContents
        .Range("IB204:IY250").ClearContents
        .Protect AllowFormattingColumns:=True, AllowFormattingRows:=True
    End With
        
    Application.ScreenUpdating = True

End Sub
