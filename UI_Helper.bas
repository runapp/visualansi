Attribute VB_Name = "UI_Helper"
Sub Text_Sel_All(ByRef ctl_text As TextBox)
    ctl_text.SelStart = 0
    ctl_text.SelLength = Len(ctl_text.text)
End Sub
