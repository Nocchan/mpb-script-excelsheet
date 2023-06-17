Attribute VB_Name = "mpb_vbascript_seasonCompletion"
Sub スペ状況リセット()

    ' エラーチェック
    If ActiveSheet.Name <> ActiveSheet.Cells(1, "A").Value & "_スケジュール" Then
        MsgBox "シート名またはA1セルのシーズン指定が不正です。"
        End
    End If

    Application.ScreenUpdating = False
    
    seasonName = ActiveSheet.Cells(1, "A")
    
    With Sheets(seasonName & "_投手データ")
        .Unprotect
        .Range("JV4:KS50").ClearContents
        .Range("JV54:KS100").ClearContents
        .Range("JV104:KS150").ClearContents
        .Range("JV154:KS200").ClearContents
        .Range("JV204:KS250").ClearContents
        .Protect AllowFormattingColumns:=True, AllowFormattingRows:=True
    End With
    
    With Sheets(seasonName & "_野手データ")
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
