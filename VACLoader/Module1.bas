Attribute VB_Name = "Module1"

Public Sub Main()
    If Dir(App.Path & "\" & "VisualAnsi.exe", vbHidden Or vbDirectory Or vbReadOnly Or vbSystem) <> "" Then
        If Shell("visualansi.exe" & " *" & Command, vbNormalFocus) = 0 Then MsgBox "�Ұʥ���", 16, "����"
        
    Else
        MsgBox "�䤣��D�{��VisualAnsi.exe !!!", 16, "�L�k�Ұ�"
    End If
    End
End Sub
