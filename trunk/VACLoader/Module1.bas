Attribute VB_Name = "Module1"

Public Sub Main()
    If Dir(App.Path & "\" & "VisualAnsi.exe", vbHidden Or vbDirectory Or vbReadOnly Or vbSystem) <> "" Then
        If Shell("visualansi.exe" & " *" & Command, vbNormalFocus) = 0 Then MsgBox "啟動失敗", 16, "失敗"
        
    Else
        MsgBox "找不到主程式VisualAnsi.exe !!!", 16, "無法啟動"
    End If
    End
End Sub
