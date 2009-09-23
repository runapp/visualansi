VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SelectLine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private SP_X As Integer
Private SP_Y As Integer
Private EP_X As Integer
Private EP_Y As Integer
Private SelectShape As Shape
Private Traget As Variant
Public Property Get StartPoint_X() As Variant
    StartPoint_X = SP_X
End Property

Public Property Let StartPoint_X(ByVal vNewValue As Variant)
    SP_X = vNewValue
    'EP_X = SP_X
    SelectShape.Left = SP_X
End Property
Public Property Get StartPoint_Y() As Variant
    StartPoint_Y = SP_Y
End Property

Public Property Let StartPoint_Y(ByVal vNewValue As Variant)
    SP_Y = vNewValue
    'EP_Y = SP_Y
    SelectShape.Top = SP_Y
End Property

Public Property Get EndPoint_X() As Variant
    EndPoint_X = EP_X
End Property

Public Property Let EndPoint_X(ByVal vNewValue As Variant)
    EP_X = vNewValue
End Property
Public Property Get EndPoint_Y() As Variant
    EndPoint_Y = EP_Y
End Property

Public Property Let EndPoint_Y(ByVal vNewValue As Variant)
    EP_Y = vNewValue
End Property

Public Property Get TragetShape() As Variant

End Property

Public Property Let TragetShape(ByVal vNewValue As Variant)
    Set SelectShape = vNewValue

End Property


Public Sub DrawSelect()
    On Error GoTo out
    'Call CorrectPoints
    Dim tempByte As Byte
    If SP_X > EP_X Then
        tempByte = 1
    Else
        tempByte = 2
    End If
    
    If SP_Y > EP_Y Then
        tempByte = tempByte
    Else
        tempByte = tempByte + 2
    End If
    DoEvents
    Select Case tempByte
        Case Is = 1
            SelectShape.Move EP_X, EP_Y, SP_X - EP_X + 1, SP_Y - EP_Y + 1
        Case Is = 2
            SelectShape.Move SP_X, EP_Y, EP_X - SP_X + 1, SP_Y - EP_Y + 1
        Case Is = 3
            SelectShape.Move EP_X, SP_Y, SP_X - EP_X + 1, EP_Y - SP_Y + 1
        Case Is = 4
            SelectShape.Move SP_X, SP_Y, EP_X - SP_X + 1, EP_Y - SP_Y + 1
        
End Select
Exit Sub
out:
Debug.Print "DrawSelect::Err:" & Err.Description
'SelectShape.Left = SP_X
'SelectShape.Top = SP_Y
'SelectShape.Width = EP_X - SP_X + 1
'SelectShape.Height = EP_Y - SP_Y + 1
'Debug.Print "(" & SP_X & "," & SP_Y; ")" & "-" & "(" & EP_X & "," & EP_Y & ")"
'SelectShape.Move SP_X, SP_Y, EP_X - SP_X + 1, EP_Y - SP_Y + 1
End Sub
Public Sub CorrectPoints()
'修正起始小於結束
    Dim tempInt As Integer
    If SP_X > EP_X Then
        tempInt = SP_X
        SP_X = EP_X
        EP_X = tempInt
    End If
    
    If SP_Y > EP_Y Then
        tempInt = SP_Y
        SP_Y = EP_Y
        EP_Y = tempInt
    End If

End Sub

Public Sub FocusStyle(ByVal isFocus As Boolean)
    If isFocus Then
        SelectShape.BorderColor = &HFF00FF
    Else
        SelectShape.BorderColor = &HFF00FF
        'SelectShape.BorderColor = &H400040
    End If
End Sub

Public Function IsInSel(ByVal x As Integer, ByVal y As Integer) As Boolean
    If x <= EP_X And x >= SP_X Then
        If y <= EP_Y And y >= SP_Y Then IsInSel = True
        
    End If
End Function
