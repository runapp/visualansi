VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HTMLCreater"
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
Public linetail As String

Public Function GetCode(ByVal Ansi As Integer, ByVal Fcolor As Byte, ByVal BColor As Byte, Optional ByVal biColor As Byte) As String
    tmpCounter = 0
    tmpFcCounter = 0
    tmp_chr = ""
    
    If Ansi = 0 Then
        Ansi = 32
        Fcolor = preFColor
    End If
    If Ansi = 60 Then
        tmp_chr = "&lt;"
    ElseIf Ansi = 62 Then
        tmp_chr = "&lt;"
    ElseIf Ansi = 34 Then
        tmp_chr = "&quot;"
    ElseIf Ansi = 38 Then
        tmp_chr = "&amp;"
    Else
        tmp_chr = Chr(Ansi)
    End If
    
    '¬Y¨ÇÂù¦ì¤¸²Õ¦r¦b ºô­¶¤U§Y¨Ï¨Ï¥Îpre¤]¥u¦³³æ¦ì¤¸¦rªº¤j¤p ¬G¥[­Óspan«ü©w¼e«×
    If (Is_Special_Char_W(Ansi)) Then
        tmp_chr = "<span class='w'>" & tmp_chr & "</span>"
    End If
    
    If preFColor = Fcolor And preBColor = BColor Then

            GetCode = tmp_chr

    Else
        Call GetCC(Fcolor, BColor)
        If tmpCounter <> 0 Then
            GetCode = "<span class='" & tmpStr & "' >" & tmp_chr
            linetail = "</span>" & linetail
        Else
            GetCode = tmp_chr
        End If
        preFColor = Fcolor
        preBColor = BColor
    End If
End Function
Public Function Is_Special_Char_W(ByVal ascval As Integer) As Boolean
    '¬Y¨ÇÂù¦ì¤¸²Õ¦r¦b ºô­¶¤U§Y¨Ï¨Ï¥Îpre¤]¥u¦³³æ¦ì¤¸¦rªº¤j¤p
    Dim tmpr As Boolean
    
    Select Case ascval
        Case -24240:
            tmpr = True
        Case Else
            tmpr = False
    End Select
    Is_Special_Char_W = tmpr
End Function
Public Function GetCode_Bi(ByVal Ansi As Integer, ByVal Fcolor As Byte, ByVal BColor As Byte, ByVal biColor As Byte) As String
    tmpCounter = 0
    tmpFcCounter = 0
    Call Clear_color
    If Ansi = 0 Then
        Ansi = 32
        Fcolor = preFColor
    End If

    If preFColor = Fcolor And preBColor = BColor Then
            If biColor = 0 Then
                GetCode_Bi = "<div class='l'>" & Chr(Ansi) & "</div>"
            Else
                GetCode_Bi = "<div class='r'>" & Chr(Ansi) & "</div>"
            End If
    Else
        Call GetCC(Fcolor, BColor, 1)
        If tmpCounter <> 0 Then
            If biColor = 0 Then
                GetCode_Bi = "<div class='" & tmpStr & " l'>" & Chr(Ansi) & "</div>"
            Else
                GetCode_Bi = "<div class='" & tmpStr & " r'>" & Chr(Ansi) & "</div>"
            End If
        Else
            If biColor = 0 Then
                GetCode_Bi = "<div class='l'>" & Chr(Ansi) & "</div>"
            Else
                GetCode_Bi = "<div class='r'>" & Chr(Ansi) & "</div>"
            End If
        End If
        'preFColor = Fcolor
        'preBColor = BColor
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
Private Sub GetCC(ByVal Fcolor As Byte, ByVal BColor As Byte, Optional biColor As Byte)

    tmpFcCounter = 0
    '³]©w«e´ºÃC¦â

    tmpColorC(1) = "c" & Fcolor
    '³B²z­I´º¦â

    tmpColorC(2) = "b" & BColor
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
            tmpStr = tmpColorC(0) & " " & tmpColorC(1)
            tmpFcCounter = tmpFcCounter + 1
        End If
        tmpCounter = tmpCounter + 1
    End If
    If preColorC(2) <> tmpColorC(2) Then
        If tmpCounter = 0 Then
            tmpStr = tmpColorC(2)
        Else
            If tmpFcCounter = 3 Then
                tmpStr = tmpStr & " " & tmpColorC(1) & " " & tmpColorC(2)
            Else
                tmpStr = tmpStr & " " & tmpColorC(2)
            End If
        End If
        tmpCounter = tmpCounter + 1
    Else
        If tmpFcCounter = 2 Then
                tmpStr = tmpStr & " " & tmpColorC(2)
        Else
            If tmpFcCounter = 3 Then
                tmpStr = tmpStr & " " & tmpColorC(1) & " " & tmpColorC(2)
            End If
        End If
    End If
    If biColor <> 1 Then
        preColorC(0) = tmpColorC(0)
        preColorC(1) = tmpColorC(1)
        preColorC(2) = tmpColorC(2)
    End If
End Sub

Public Sub Clear()
'¦^ÂkÃC¦âªº¹w³]­È
preColorC(0) = ""
preColorC(1) = ""
preColorC(2) = ""
preFColor = 7
preBColor = 0
'preColorC(0) = ""
'preColorC(1) = "c7"
'preColorC(2) = "b0"
'preFColor = 7
'preBColor = 0
linetail = "</div>"
End Sub

Public Sub Clear_color()
'¦^ÂkÃC¦âªº¹w³]­È
preColorC(0) = ""
preColorC(1) = ""
preColorC(2) = ""
preFColor = 7
preBColor = 0

End Sub
Private Sub Class_Initialize()
'³]©wÃC¦âªº¹w³]­È
'preColorC(0) = ""
'preColorC(1) = "c7"
'preColorC(2) = "b0"
'preFColor = 7
'preBColor = 0
preColorC(0) = ""
preColorC(1) = ""
preColorC(2) = ""
preFColor = 7
preBColor = 0
End Sub


