Attribute VB_Name = "WINAPI_Msg"
Option Explicit
Public Const GWL_WNDPROC = (-4)
Public Const WM_NCHITTEST = &H84
Public Const WM_NCMOUSEMOVE = &HA0
Public Const WM_NCLBUTTONDBLCLK = &HA3
Public Const WM_NCLBUTTONDOWN = &HA1
Public Const WM_NCLBUTTONUP = &HA2
Public Const WM_NCMBUTTONDBLCLK = &HA9
Public Const WM_NCMBUTTONDOWN = &HA7
Public Const WM_NCMBUTTONUP = &HA8
Public Const WM_NCRBUTTONDBLCLK = &HA6
Public Const WM_NCRBUTTONDOWN = &HA4
Public Const WM_NCRBUTTONUP = &HA5
Public Const WM_KEYDOWN = &H100


Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Public prevWndProc As Long
Public prevtext2WndProc As Long

Function WndProc(ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim sMsg As String
    If Msg = WM_NCLBUTTONDOWN Then Debug.Print "mouse down"

    'Debug.Print Msg
    WndProc = CallWindowProc(prevWndProc, hwnd, Msg, wParam, lParam)
    
    If Msg = 3 Then
        Call Form14.CDrawer.ReFreshColor
        Debug.Print "ReFreshColor"
    End If
End Function
Function WndProc_Text2(ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    If Msg = WM_KEYDOWN Then
        Debug.Print "TEXT2::WM_KEYDOWN wParam=" & wParam
    End If
End Function



