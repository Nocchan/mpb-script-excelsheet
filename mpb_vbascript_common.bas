Attribute VB_Name = "mpb_vbascript_common"
Sub バックアップ()
    
    ActiveWorkbook.SaveCopyAs Filename:="C:\Users\TaiNo\Desktop\ペナントバックアップ\" & Format(Now, "yyyymmddhhnnss") & ".xlsm"
    
End Sub

