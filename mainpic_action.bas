Attribute VB_Name = "Msg_Pic"
Public prevWndProc_Pic As Long
Function WndProc_Pic(ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim sMsg As String
    If Msg = WM_NCLBUTTONDOWN Then
        Form1.Pic1MouseDown = True
    End If
    If Msg = WM_NCLBUTTONUP Then
        Form1.Pic1MouseDown = False
    End If
    'Debug.Print Msg
    WndProc_Pic = CallWindowProc(prevWndProc_Pic, hwnd, Msg, wParam, lParam)
    
End Function
