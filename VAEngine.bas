Attribute VB_Name = "VAENGINE"
Public Sub DoDraw_A(ByRef Arr() As ColorLayer, ByVal x As Integer, ByVal y As Integer, ByVal z As Integer, ByVal tstr As String, ByVal Fcolor As Byte)
    On Error GoTo out

    If Asc(tstr) = 0 Or Asc(tstr) = 32 Then Exit Sub
    If Tlen(tstr) = 2 Then
        If x >= UBound(Arr, 1) Then Exit Sub
        If Arr(x, y, z).Ansi <> 0 Or Arr(x + 1, y, z).Ansi <> 0 Then Exit Sub
    
    Else
        If Arr(x, y, z).Ansi <> 0 Then Exit Sub
    End If

    Arr(x, y, z).Ansi = Asc(tstr)
    Arr(x, y, z).Color = Fcolor
    If Tlen(tstr) = 2 Then
        Arr(x + 1, y, z).Ansi = -1
        Arr(x + 1, y, z).Color = Fcolor
        'Arr(X + 1, Y, Z).BColor = Arr(X, Y, Z).BColor
    End If

Exit Sub
out:
Debug.Print "DoDrawA Error Out : " & Err.Description
End Sub

Public Sub DoDrawBC_A(ByRef Arr() As ColorLayer, ByVal x As Integer, ByVal y As Integer, ByVal z As Integer, ByVal BColor As Byte)
On Error GoTo out
If Arr(x, y, z).BColor = 0 Then
    If Arr(x, y, z).Ansi = -1 Then
        Arr(x - 1, y, z).BColor = BColor
    End If
    Arr(x, y, z).BColor = BColor
    If x <> UBound(Arr, 1) Then
        If Arr(x + 1, y, z).Ansi = -1 Then
            Arr(x + 1, y, z).BColor = BColor
        End If
    End If
End If
Exit Sub
out:
Debug.Print "DoDrawBC Error Out" & Err.Description


End Sub
Public Function DoErease_A(ByRef Arr() As ColorLayer, ByVal x As Integer, ByVal y As Integer, ByVal z As Integer) As Integer
On Error GoTo out

'�^�ǧR�������t�q(���줸�r���B�z)
'+1���A��ܮɻ����x+1������
' 0����줸�r
'-1���A��ܮɻ����x-1������
If Arr(x, y, z).Ansi <> 0 Then
    
    If Arr(x, y, z).Ansi = -1 Then
        Arr(x - 1, y, z).Ansi = 0
        DoErease_A = -1
    End If
    Arr(x, y, z).Ansi = 0
    If Arr(x + 1, y, z).Ansi = -1 Then
        Arr(x + 1, y, z).Ansi = 0
        DoErease_A = 1
    End If
End If

Exit Function
out:
Debug.Print "DoErease_A Error Out" & Err.Description
End Function


Public Function DoEreaseB_A(ByRef Arr() As ColorLayer, ByVal x As Integer, ByVal y As Integer, ByVal z As Integer) As Integer
On Error GoTo out
'�^�ǧR�������t�q(���줸�r���B�z)
'+1���A��ܮɻ����x+1������
' 0����줸�r
'-1���A��ܮɻ����x-1������
If Arr(x, y, z).BColor <> 0 Then
    
    If Arr(x, y, z).Ansi = -1 Then
        Arr(x - 1, y, z).BColor = 0
        DoEreaseB_A = -1
    End If
    Arr(x, y, z).BColor = 0
    
    If Arr(x + 1, y, z).Ansi = -1 Then
        Arr(x + 1, y, z).BColor = 0
        DoEreaseB_A = 1

    End If

End If
Exit Function
out:
Debug.Print "DoEreaseB_A Error Out" & Err.Description
End Function


Public Sub CLArrayCopy(ByRef fromArr() As ColorLayer, ByRef newArr() As ColorLayer, ByVal x1 As Integer, ByVal Y1 As Integer, ByVal X2 As Integer, ByVal Y2 As Integer, ByVal z As Integer)
    On Error GoTo out
    Dim ubX As Integer, ubY As Integer
    ubX = Abs(X2 - x1)
    ubY = Abs(Y2 - Y1)
    ReDim newArr(ubX, ubY)
    For j = 0 To ubY
        For i = 0 To ubX
            If i = ubX Then
                If fromArr(x1 + i + 1, Y1 + j, z).Ansi <> -1 Then newArr(i, j) = fromArr(x1 + i, Y1 + j, z)
            ElseIf i = 0 Then
                If fromArr(x1 + i, Y1 + j, z).Ansi <> -1 Then newArr(i, j) = fromArr(x1 + i, Y1 + j, z)
            Else
                newArr(i, j) = fromArr(x1 + i, Y1 + j, z)
            End If
        Next i
    Next j
    Exit Sub
out:
    Debug.Print "CLArrayCopy::Err:"; Err.Description
End Sub

Public Sub CLArrayPaste_C(ByRef fromArr() As ColorLayer, ByRef toArr() As ColorLayer, ByVal x As Integer, ByVal y As Integer, ByVal z As Integer)
    'fromArr�O2D toArr�O3D
    '���A�K�W���ɫ�h�I
    Dim fromUbX As Integer, toUbX As Integer, fromUbY As Integer, toUbY As Integer
    Dim tmpAnsiStr As String
    Dim tmpFColor As Byte
    fromUbX = UBound(fromArr, 1)
    fromUbY = UBound(fromArr, 2)
    toUbX = UBound(toArr, 1)
    toUbY = UBound(toArr, 2)
    
    For j = 0 To fromUbY
        For i = 0 To fromUbX
            If fromArr(i, j).Ansi > 256 Or fromArr(i, j).Ansi < 0 Then  '���줸�r
                If x + i + 1 <= toUbX And i <= fromUbX Then
                    If fromArr(i + 1, j).Ansi = -1 Then
                        Call DoErease_A(toArr, x + i, y + j, z)
                        Call DoErease_A(toArr, x + 1 + i, y + j, z)
                        toArr(x + i, y + j, z) = fromArr(i, j)
                        toArr(x + i + 1, y + j, z) = fromArr(i + 1, j)
                        i = i + 1
                    End If
                End If
            'ElseIf (fromArr(i, j).Ansi = 0 Or fromArr(i, j).Ansi <> 32) And fromArr(i, j).BColor = 0 Then
                
            Else
            
                If fromArr(i, j).Ansi <> 0 And fromArr(i, j).Ansi <> 32 Or fromArr(i, j).BColor <> 0 Then '�T�{�n�g�J�����e���O�ť�
                    '�}�lŪ���g�J�����e
                    tmpAnsiStr = Chr(fromArr(i, j).Ansi)
                    '�g�J
                    Call DoErease_A(toArr, x + i, y + j, z)
                    Call DoDraw_A(toArr, x + i, y + j, z, tmpAnsiStr, fromArr(i, j).Color)
                    
                End If
                If fromArr(i, j).BColor <> 0 Then   '�T�{�n�g�J���C�⤣�O�¦�
                    Call DoEreaseB_A(toArr, x + i, y + j, z)
                    Call DoDrawBC_A(toArr, x + i, y + j, z, fromArr(i, j).BColor)
                End If
            End If
            DoEvents
        Next i
    Next j
    
End Sub
Public Sub PaintColor_A_bak(ByRef Arr() As ColorLayer, ByVal x As Integer, ByVal y As Integer, ByVal z As Integer, ByVal Fcolor As Byte, ByVal BColor As Byte, Optional ByVal F As Byte, Optional ByVal B As Byte)
On Error GoTo out

    If F = 1 Then
        If Arr(x, y, z).Ansi = -1 Then
            Arr(x - 1, y, z).Color = Fcolor
        End If
        Arr(x, y, z).Color = Fcolor
        If Arr(x + 1, y, z).Ansi = -1 Then
            Arr(x + 1, y, z).Color = Fcolor
        End If
        If Arr(x, y, z).Ansi = 0 Or Arr(x, y, z).Ansi = 32 Then
            Arr(x, y, z).Color = 7
        End If
    End If
    
    If B = 1 Then
        If Arr(x, y, z).Ansi = -1 Then
            Arr(x - 1, y, z).BColor = BColor
        End If
        Arr(x, y, z).BColor = BColor
        If Arr(x + 1, y, z).Ansi = -1 Then
            Arr(x + 1, y, z).BColor = BColor
        End If
    End If
    
Exit Sub
out:
    Debug.Print "Paint Color Error Out"
    Resume Next
End Sub
Public Sub PaintColor_A(ByRef Arr() As ColorLayer, ByVal x As Integer, ByVal y As Integer, ByVal z As Integer, ByVal Fcolor As Byte, ByVal BColor As Byte, Optional ByVal F As Byte, Optional ByVal B As Byte)
On Error GoTo out
    '����r����
    If F = 1 Then
        Arr(x, y, z).Color = Fcolor
        If Arr(x, y, z).Ansi = 0 Or Arr(x, y, z).Ansi = 32 Then
            Arr(x, y, z).Color = 7
        End If
    End If
    
    If B = 1 Then
        Arr(x, y, z).BColor = BColor
    End If
    
Exit Sub
out:
    Debug.Print "Paint Color Error Out"
    Resume Next
End Sub
Public Sub PaintColor_B(ByRef Arr() As ColorLayer, ByVal x As Integer, ByVal y As Integer, ByVal z As Integer, ByVal Fcolor As Byte, ByVal BColor As Byte, Optional ByVal F As Byte, Optional ByVal B As Byte)
On Error GoTo out

    If F = 1 Then
        
        If Arr(x, y, z).Ansi = 0 Or Arr(x, y, z).Ansi = 32 Then
            Arr(x, y, z).Color = 7
        Else
            Arr(x, y, z).Color = Fcolor
        
        End If
    End If
    
    If B = 1 Then
        Arr(x, y, z).BColor = BColor
    End If
    
Exit Sub
out:
    Debug.Print "Paint Color Error Out"
    Resume Next
End Sub

Public Sub ExChColor_Draw_A_BAK(ByRef Arr() As ColorLayer, ByVal x As Integer, ByVal y As Integer, ByVal z As Integer, ByVal preFColor As Byte, ByVal preBColor As Byte, ByVal newFColor As Byte, ByVal newBColor As Byte, Optional ByVal F As Byte, Optional ByVal B As Byte)
On Error GoTo out
    'Debug.Print "preFColor=" & preFColor & vbCrLf & "Arrf(X, Y, z).Color=" & arr(X, Y, z).Color

    If F = 1 And Arr(x, y, z).Color = preFColor Then
        If Arr(x, y, z).Ansi = -1 Then
            Arr(x - 1, y, z).Color = newFColor
        End If
        Arr(x, y, z).Color = newFColor
        If Arr(x + 1, y, z).Ansi = -1 Then
            Arr(x + 1, y, z).Color = newFColor
        End If
        If Arr(x, y, z).Ansi = 0 Or Arr(x, y, z).Ansi = 32 Then
            Arr(x, y, z).Color = 7
        End If
    End If
    
    If B = 1 And Arr(x, y, z).BColor = preBColor Then
        If Arr(x, y, z).Ansi = -1 Then
            Arr(x - 1, y, z).BColor = newBColor
        End If
        Arr(x, y, z).BColor = newBColor
        If Arr(x + 1, y, z).Ansi = -1 Then
            Arr(x + 1, y, z).BColor = newBColor
        End If
    End If
Exit Sub
out:

    Debug.Print "ExChColor_Draw_A  Error Out"
    Resume Next
End Sub

Public Sub ExChColor_Draw_A(ByRef Arr() As ColorLayer, ByVal x As Integer, ByVal y As Integer, ByVal z As Integer, ByVal preFColor As Byte, ByVal preBColor As Byte, ByVal newFColor As Byte, ByVal newBColor As Byte, Optional ByVal F As Byte, Optional ByVal B As Byte)
On Error GoTo out
    'Debug.Print "preFColor=" & preFColor & vbCrLf & "Arrf(X, Y, z).Color=" & arr(X, Y, z).Color

    If F = 1 And Arr(x, y, z).Color = preFColor Then

        Arr(x, y, z).Color = newFColor
        If Arr(x, y, z).Ansi = 0 Or Arr(x, y, z).Ansi = 32 Then
            Arr(x, y, z).Color = 7
        End If
    End If
    
    If B = 1 And Arr(x, y, z).BColor = preBColor Then
        Arr(x, y, z).BColor = newBColor
    End If
Exit Sub
out:

    Debug.Print "ExChColor_Draw_A  Error Out"
    Resume Next
End Sub


Public Sub Eff_Move_Area(ByVal x1 As Integer, ByVal Y1 As Integer, ByVal X2 As Integer, ByVal Y2 As Integer, ByVal z As Integer, ByVal Xshift As Integer, ByVal Yshift As Integer)
    Dim exCheck As Byte
    Dim i As Integer, j As Integer, maxX As Integer, fromX As Integer, toX As Integer, fromY As Integer, toY As Integer, tmpInt As Integer, iStep As Integer, jStep As Integer
    Dim tmpCC As ColorLayer
    maxX = UBound(Arrf, 1)

    '�j�p���ǭȰ���
    If X2 < x1 Then
        tmpInt = x1
        x1 = X2
        X2 = tmpInt
    End If
    If Y2 < Y1 Then
        tmpInt = x1
        Y1 = Y2
        Y2 = tmpInt
    End If
    '�M�w��V
    If Xshift > 0 Then
        iStep = -1
        fromX = X2
        toX = x1
    Else
        iStep = 1
        fromX = x1
        toX = X2
    End If
    If Yshift > 0 Then
        jStep = -1
        fromY = Y2
        toY = Y1
    Else
        jStep = 1
        fromY = Y1
        toY = Y2
    End If
    
    For j = fromY To toY Step jStep
    
        For i = fromX To toX Step iStep
            'If i > maxX Then
            '    Call DoErease_A(Arrf, i + Xshift, j + Yshift, Z)
            '    Arrf(i + Xshift, j + Yshift, Z) = Arrf(i, j, Z) '�ƻs�r��
            '    Arrf(i, j, Z) = tmpCC   '�R���쥻��
            'ElseIf i = toX Then
            'ElseIf i = fromX Then
            If i <> maxX Then
                If Arrf(i, j, z).Ansi = -1 Then
                    '�B�z�Ĥ@�ӹJ��-1
                    '���k �G�q�k�쥪
                    If i = fromX And Xshift < 0 Then
                        '�S�� ���� ���̥��r��
                        Call DoErease_A(Arrf, i + Xshift - 1, j + Yshift, z)
                        Arrf(i + Xshift - 1, j + Yshift, z) = Arrf(i - 1, j, z) '�ƻs�r��
                        
                        Arrf(i - 1, j, z) = tmpCC
                        Call DoErease_A(Arrf, i + Xshift, j + Yshift, z)
                        Arrf(i + Xshift, j + Yshift, z) = Arrf(i, j, z)
                        
                        Arrf(i, j, z) = tmpCC
                        '�h���ʤ@��
                        i = i + iStep
                    Else
                        Call DoErease_A(Arrf, i + Xshift, j + Yshift, z)
                        Arrf(i + Xshift, j + Yshift, z) = Arrf(i, j, z) '�ƻs�r��
                        Arrf(i, j, z) = tmpCC
                        Call DoErease_A(Arrf, i + Xshift - 1, j + Yshift, z)
                        Arrf(i + Xshift - 1, j + Yshift, z) = Arrf(i - 1, j, z)
                        
                        Arrf(i - 1, j, z) = tmpCC
                        '�h���ʤ@��
                        i = i + iStep
                    End If
                ElseIf Arrf(i + 1, j, z).Ansi = -1 Then
                    '�B�z�᭱���ȬO-1
                    '������ �G�q����k
                    If i = fromX And Xshift > 0 Then
                        '�S�� ���k ���̥k�r��
                        Call DoErease_A(Arrf, i + Xshift + 1, j + Yshift, z)
                        Arrf(i + Xshift + 1, j + Yshift, z) = Arrf(i + 1, j, z) '�ƻs�r��
                        Arrf(i + 1, j, z) = tmpCC
                        Call DoErease_A(Arrf, i + Xshift, j + Yshift, z)
                        Arrf(i + Xshift, j + Yshift, z) = Arrf(i, j, z)
                        
                        Arrf(i, j, z) = tmpCC
                        '�h���ʤ@��
                        i = i + iStep

                    Else
                        Call DoErease_A(Arrf, i + Xshift, j + Yshift, z)
                        Arrf(i + Xshift, j + Yshift, z) = Arrf(i, j, z) '�ƻs�r��
                        
                        Arrf(i, j, z) = tmpCC
                        Call DoErease_A(Arrf, i + Xshift + 1, j + Yshift, z)
                        Arrf(i + Xshift + 1, j + Yshift, z) = Arrf(i + 1, j, z)
                        
                        Arrf(i + 1, j, z) = tmpCC
                        '�h���ʤ@��
                        i = i + iStep
                    End If
                Else
                    
                    Call DoErease_A(Arrf, i + Xshift, j + Yshift, z)
                    Arrf(i + Xshift, j + Yshift, z) = Arrf(i, j, z) '�ƻs�r��
                    Arrf(i, j, z) = tmpCC   '�R���쥻��
                End If
            Else
                If Arrf(i, j, z).Ansi = -1 Then
                    '�B�z�Ĥ@�ӹJ��-1
                    '���k �G�q�k�쥪
                    If i = toX And Xshift <= 0 Then
                        Call DoErease_A(Arrf, i + Xshift - 1, j + Yshift, z)
                        Arrf(i + Xshift - 1, j + Yshift, z) = Arrf(i - 1, j, z) '�ƻs�r��
                        
                        Arrf(i - 1, j, z) = tmpCC
                        Call DoErease_A(Arrf, i + Xshift, j + Yshift, z)
                        Arrf(i + Xshift, j + Yshift, z) = Arrf(i, j, z)
                        
                        Arrf(i, j, z) = tmpCC
                        '�h���ʤ@��
                        i = i + iStep
                    Else
                        Call DoErease_A(Arrf, i + Xshift, j + Yshift, z)
                        Arrf(i + Xshift, j + Yshift, z) = Arrf(i, j, z) '�ƻs�r��
                        Arrf(i, j, z) = tmpCC
                        Call DoErease_A(Arrf, i + Xshift - 1, j + Yshift, z)
                        Arrf(i + Xshift - 1, j + Yshift, z) = Arrf(i - 1, j, z)
                        
                        Arrf(i - 1, j, z) = tmpCC
                        '�h���ʤ@��
                        i = i + iStep
                    End If
    
                Else
                    Call DoErease_A(Arrf, i + Xshift, j + Yshift, z)
                    Arrf(i + Xshift, j + Yshift, z) = Arrf(i, j, z) '�ƻs�r��
                    Arrf(i, j, z) = tmpCC   '�R���쥻��
                End If
            End If
        Next i
    
    Next j
    
    
End Sub
Public Function Get_Char(x As Integer, y As Integer, z As Integer) As String
    '�^�ǸӮy�Цr��
    If Arrf(x, y, OFP.CurrentPage).Ansi = -1 Then
        If x <> 0 Then
            Get_Char = Chr(Arrf(x - 1, y, OFP.CurrentPage).Ansi)
        End If
    Else
        Get_Char = Chr(Arrf(x, y, OFP.CurrentPage).Ansi)
    End If
End Function
