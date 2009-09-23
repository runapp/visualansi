Attribute VB_Name = "HotKeyModule"
Sub HotKey_Handler(ByVal KeyCode As Integer, ByVal Shift As Integer)
      Debug.Print "hotkey:" & KeyCode & "(" & Shift & ")"
      
      Select Case KeyCode
        Case 83:     'Ctrl+S Save
            If OFP.Closed = False And Shift = 2 Then Form1.Me_Save_Click
        Case 78:     'Ctrl+N New
            If Shift = 2 Then Form1.Me_New_Click
        Case 79:     'Ctrl+O Open
            If Shift = 2 Then Form1.Me_OpenFile_Click
        Case 34:    'pagedown
            If OFP.Closed = False And Form1.Command7(1).Enabled Then
                Form1.Command7_Click (1)
            End If
        Case 33:    'pageUp
            If OFP.Closed = False And Form1.Command7(0).Enabled Then
                Form1.Command7_Click (0)
            End If
        Case Else
            
      End Select
      
      
End Sub
