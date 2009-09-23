VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ColorCodeCreater"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private preColorC(2) As String
Private tmpColorC(2) As String
Public preFColor As Byte
Public preBColor As Byte
Private tmpStr As String
Private tmpCounter As Byte
Private tmpFcCounter As Byte

Public Function GetCode(ByVal Ansi As Integer, ByVal Fcolor As Byte, ByVal BColor As Byte) As String
    tmpCounter = 0
    tmpFcCounter = 0
    If Ansi = 0 Then
        Ansi = 32
        Fcolor = preFColor
    End If
    If preFColor = Fcolor And preBColor = BColor Then
        GetCode = Chr(Ansi)
    Else
        Call GetCC(Fcolor, BColor)
        If tmpCounter <> 0 Then
            GetCode = "[" & tmpStr & "m" & Chr(Ansi)
        Else
            GetCode = Chr(Ansi)
        End If
        preFColor = Fcolor
        preBColor = BColor
    End If
End Function
Public Function GetCode_noChar(ByVal Fcolor As Byte, ByVal BColor As Byte) As String
    tmpCounter = 0
    tmpFcCounter = 0
    If preFColor = Fcolor And preBColor = BColor Then
        GetCode_noChar = ""
    Else
        Call GetCC(Fcolor, BColor)
        If tmpCounter <> 0 Then
            GetCode_noChar = "[" & tmpStr & "m"
        Else
            GetCode_noChar = ""
        End If
        preFColor = Fcolor
        preBColor = BColor
    End If
End Function
Private Sub GetCC(ByVal Fcolor As Byte, ByVal BColor As Byte)
    '³]©w«e´º«G

    If Fcolor \ 8 = 1 Then
        tmpColorC(0) = "1"
    Else
        tmpColorC(0) = ""
        tmpFcCounter = 1
    End If
    '³]©w«e´ºÃC¦â
    Select Case Fcolor Mod 8
        Case Is = 0
            tmpColorC(1) = "30"
        Case Is = 1
            tmpColorC(1) = "34"
        Case Is = 2
            tmpColorC(1) = "32"
        Case Is = 3
            tmpColorC(1) = "36"
        Case Is = 4
            tmpColorC(1) = "31"
        Case Is = 5
            tmpColorC(1) = "35"
        Case Is = 6
            tmpColorC(1) = "33"
        Case Is = 7
            tmpColorC(1) = "37"
    End Select
    '³B²z­I´º¦â
    Select Case BColor Mod 8
        Case Is = 0
            tmpColorC(2) = "40"
        Case Is = 1
            tmpColorC(2) = "44"
        Case Is = 2
            tmpColorC(2) = "42"
        Case Is = 3
            tmpColorC(2) = "46"
        Case Is = 4
            tmpColorC(2) = "41"
        Case Is = 5
            tmpColorC(2) = "45"
        Case Is = 6
            tmpColorC(2) = "43"
        Case Is = 7
            tmpColorC(2) = "47"
    End Select
    '²£¥Í¦r¦ê
    tmpStr = ""
    If preColorC(1) = tmpColorC(1) Then
        If preColorC(0) <> tmpColorC(0) Then
            tmpStr = tmpColorC(0)
            tmpCounter = tmpCounter + 1
            tmpFcCounter = tmpFcCounter + 2
        End If
    Else
        If preColorC(0) = tmpColorC(0) Then
            tmpStr = tmpColorC(1)
        Else
            tmpStr = tmpColorC(0) & ";" & tmpColorC(1)
            tmpFcCounter = tmpFcCounter + 1
        End If
        tmpCounter = tmpCounter + 1
    End If
    If preColorC(2) <> tmpColorC(2) Then
        If tmpCounter = 0 Then
            tmpStr = tmpColorC(2)
        Else
            If tmpFcCounter = 3 Then
                tmpStr = tmpStr & ";" & tmpColorC(1) & ";" & tmpColorC(2)
            Else
                tmpStr = tmpStr & ";" & tmpColorC(2)
            End If
        End If
        tmpCounter = tmpCounter + 1
    Else
        If tmpFcCounter = 2 Then
                tmpStr = tmpStr & ";" & tmpColorC(2)
        Else
            If tmpFcCounter = 3 Then
                tmpStr = tmpStr & ";" & tmpColorC(1) & ";" & tmpColorC(2)
            End If
        End If
    End If
    preColorC(0) = tmpColorC(0)
    preColorC(1) = tmpColorC(1)
    preColorC(2) = tmpColorC(2)
End Sub

Public Sub Clear()
'¦^ÂkÃC¦âªº¹w³]­È
preColorC(0) = ""
preColorC(1) = "37"
preColorC(2) = "40"
preFColor = 7
preBColor = 0
End Sub

Private Sub Class_Initialize()
'³]©wÃC¦âªº¹w³]­È
preColorC(0) = ""
preColorC(1) = "37"
preColorC(2) = "40"
preFColor = 7
preBColor = 0
End Sub

