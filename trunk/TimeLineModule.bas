Attribute VB_Name = "TimeLineModule"
 
Public Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Public Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long
Private Const ID_Timer = 1
Public TimerFlag As Byte
Public DelayFlag As Byte
Public Sub SetTimeLine(ByVal fromPage As Integer, ByVal toPage As Integer, ByVal Value As Single)
    Dim i As Integer
    For i = fromPage To toPage
        timeLine(i) = Value
    Next i

End Sub

Public Sub SetTimerTime(ByVal hwnd As Long, ByVal intime As Long)
    SetTimer hwnd, ID_Timer, intime, AddressOf TimerFunc
End Sub

Public Sub KillTimerEnd(ByVal hwnd As Long)
    KillTimer hwnd, ID_Timer
End Sub

Sub TimerFunc(ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long)
    Call Form1.TimerFunc
    Debug.Print "timer.run"
End Sub

Sub SensDelay(ByVal hwnd As Long, ByVal intime As Long)
    On Error GoTo out
    SetTimer hwnd, 2, intime, AddressOf SensDelayFunc
Exit Sub
out:
Debug.Print "SensDelay::Err:" & Err.Description
End Sub
Sub SensDelayFunc(ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long)
    On Error GoTo out
    DelayFlag = 2
    Call Form1.Text2_Change
    DelayFlag = 1
    KillTimer hwnd, 2
Exit Sub
out:
Debug.Print "SensDelayFunc::Err:" & Err.Description
End Sub

Sub CopyDelay(ByVal hwnd As Long, ByVal intime As Long)
    On Error GoTo out
    SetTimer hwnd, 3, intime, AddressOf CopyDelayFunc
Exit Sub
out:
Debug.Print "SensDelay::Err:" & Err.Description
End Sub
Sub CopyDelayFunc(ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long)
    KillTimer hwnd, 3
    Call Form1.Command23_Click
End Sub
