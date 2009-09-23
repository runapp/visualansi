VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ColorCodeReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private tmpCCStr As String 'Temp Color Code String
Private tmpCLA() As ColorLayer
Private tmpVArr() As Variant
Private Ba_CLA() As ColorLayer


Public Property Get SetCCStr() As Variant
    SetCCStr = tmpCCStr
End Property

Public Property Let SetCCStr(ByVal vNewValue As Variant)
    tmpCCStr = vNewValue
    
End Property

Public Sub AnalyzeCC()
On Error GoTo out
'¥D°}¦CtmpVArr(¦æ¯Á¤Þ)(°Ï¬q¯Á¤Þ,0 or 1) 0¦s¦â½X°}¦C 1¦s¦r¦ê
'«Å§iÅÜ¼Æ
    Dim tmpStrLA() As String    'Àx¦s¨C¦æ
    Dim tmpstrA() As String     'Àx¦s¦U­Ó©R¥O°Ï¬q
    Dim tmpStrA1() As String    'Àx¦s0©R¥O1¦r¦ê
    Dim tmpInt As Integer       '¯Âºé
    Dim IntCounter As Integer
    Dim tmpIntCounter As Integer
    Dim tmpbyte As Byte
    Dim tmpStr As String * 1
'²Ä¤@¶¥¬q«Ø¥ßSTR¥D°}¦C
    tmpStrLA = Split(tmpCCStr, vbCrLf) '±N¦æ¤À¶}
    tmpInt = UBound(tmpStrLA)   'tmpint¬°¦æ¼Æ-1
    ReDim tmpVArr(tmpInt)    '«Ø¥ß¥D°}¦C
    For i = 0 To tmpInt
        If tmpStrLA(i) <> "" Then
            tmpstrA = Split(tmpStrLA(i), "[")  '«ü¥O°Ï¬q¤ÀÂ÷
            ReDim tmpStrA1(UBound(tmpstrA), 1)  '²Ä¤Gºûªº 0¦s¦â½X°}¦C 1¦s¦r¦ê
            tmpIntCounter = 0  '±N­pºâ¦r¼ÆªºcounterÂk¹s
            For j = 0 To UBound(tmpstrA)
                tmpInt = InStr(tmpstrA(j), "m")
                If tmpInt <> 0 Then
                    tmpStrA1(j, 0) = Replace(Left(tmpstrA(j), tmpInt - 1), " ", "") 'ÃD¨ú¥X¦â±m½X ¨Ã¥h°£©Ò¦³ªÅ¥Õ
                End If
                tmpStrA1(j, 1) = Right(tmpstrA(j), Len(tmpstrA(j)) - tmpInt)    'ÃD¨ú¥XANSI¦r¦ê³¡¤À
                tmpIntCounter = tmpIntCounter + Tlen(tmpStrA1(j, 1))    '±N¦r¼Æ²Ö­p¨ìCOUNTER¤¤
                tmpVArr(i) = tmpStrA1   '±N¦¹¦æ¤§°}¦C©ñ¤J¥D°}¦C
                
            Next j
        Else
            ReDim tmpStrA1(0, 1)
            tmpVArr(i) = tmpStrA1
        End If
        If tmpIntCounter > IntCounter Then IntCounter = tmpIntCounter   '¦pªG¦¹¦æ¦r¼Æªº¤ñ¸û¦h´N³]©w¼e¬°¦¹¦æ¦æ¼Æ
    Next i
    If IntCounter = 0 Then
        Debug.Print "error ¶×¤J±m¦â½X¶K¤Wªº¼e«×¬O0  "
        Exit Sub
    End If
'²Ä¤G¶¥¬q±NSTR°}¦C»s¦¨COLOR LAYER ARRAY
    Dim tmpStrA3() As String '¤ÀªR¦â½Xªº¼È¦s®e¾¹
    Dim tmpCountPerLine As Integer  '¦æ¤¤¦r¼Æ«ü¼Ð(tlen)
    Dim tmpCL As ColorLayer
    Dim tmpCLbyte(2) As Byte
    If IntCounter = 0 Then Debug.Print "¿ù»~: ¼e«×¤£¥i¬°¹s"
    ReDim tmpCLA(IntCounter - 1, UBound(tmpVArr)) '«Ø¥ßCOLOR LAYER°}¦C
    Debug.Print "³Ì¼e¼e«×: " & IntCounter
    tmpCL.BColor = 0
    tmpCL.Color = 7
    For i = 0 To UBound(tmpVArr)
        tmpCountPerLine = 1 '±N«ü¼ÐÂk¹s
        For j = 0 To UBound(tmpVArr(i), 1)
            'Debug.Print "Line:" & i; " ;Section: " & j & ";Len=" & Tlen(tmpVArr(i)(j, 1)) & "Pos:" & tmpCountPerLine
            If tmpVArr(i)(j, 0) = "" And j <> 0 Then
                '¦â±m¥þ³¡ÁÙ­ì¬°¹w³]­È
                tmpCL.BColor = 0
                tmpCL.Color = 7
            Else
                tmpStrA3 = Split(tmpVArr(i)(j, 0), ";") '±N¦U¦â½X¤ÀÂ÷
                If UBound(tmpStrA3) = 0 Then
                '·í¦â½X¥u¦³¤@­Ó®É
                    tmpbyte = Val(tmpStrA3(0))
                    Select Case tmpbyte
                        Case 1
                            If tmpCL.Color \ 8 = 0 Then tmpCL.Color = tmpCL.Color + 8
                        Case Is >= 40
                            tmpCL.BColor = GetBC(tmpbyte)
                        Case Is >= 30
                            tmpCL.Color = GetFC(tmpbyte) + (tmpCL.Color \ 8) * 8
                    End Select
                Else
                    'Debug.Print "¦h¦â½X:" & UBound(tmpStrA3)
                    tmpCLbyte(0) = 0
                    tmpCLbyte(1) = 255
                    tmpCLbyte(2) = 255
                    For k = 0 To UBound(tmpStrA3)
                    '¥ý³B²z¦â±m
                        tmpbyte = Val(tmpStrA3(k))
                        '³]©w¼È¦s¤§¼È¦s¹w³] 255ªº¼Æ¦r¥u¬O¬°¤F¸ò0§@¦³¨S¦³ÅÜ§óªº¤ÀÅG¥Î

                        Select Case tmpbyte
                            'ÁÙ­ì¦â±m¼È¦s¬°¹w³]
                            Case 0
                                tmpCL.BColor = 0
                                tmpCL.Color = 7
                                'Debug.Print "¼È¦s¦â±mÁÙ­ì¹w³]"
                            '¹Bºâ¦â±m¼È¦s¤§¼È¦s ¥Î©ó¤ñ¸û
                            Case Is = 1
                                tmpCLbyte(0) = 8
                            Case Is >= 40
                                tmpCLbyte(2) = GetBC(tmpbyte)
                            Case Is >= 30
                                tmpCLbyte(1) = GetFC(tmpbyte)
                        End Select
                    Next k
                    '±q(¦â±m¼È¦s¤§¼È¦s)¼g¤J(¦â±m¼È¦s)
                    If tmpCLbyte(1) <> 255 Then
                        tmpCL.Color = tmpCLbyte(1) + (tmpCL.Color \ 8) * 8
                    End If
                    If tmpCLbyte(0) <> 0 And tmpCL.Color \ 8 = 0 Then
                        tmpCL.Color = tmpCL.Color + tmpCLbyte(0)
                        'Debug.Print "«G¦â"
                    End If
                    If tmpCLbyte(2) <> 255 Then
                        tmpCL.BColor = tmpCLbyte(2)
                    End If
                    
                End If

            End If
                    '¦A¼g¤J¤å¦r
                    If tmpVArr(i)(j, 1) <> "" Then
                        For L = 1 To Len(tmpVArr(i)(j, 1))
                            tmpStr = Mid(tmpVArr(i)(j, 1), L, 1)
                            'Debug.Print "Line:" & i; " ;section: " & j & ";¦r¼Æ:" & tmpCountPerLine - 1
                            tmpCLA(tmpCountPerLine - 1, i).Ansi = Asc(tmpStr)
                            tmpCLA(tmpCountPerLine - 1, i).Color = tmpCL.Color
                            tmpCLA(tmpCountPerLine - 1, i).BColor = tmpCL.BColor
                            tmpbyte = Tlen(tmpStr)
                            If tmpbyte = 2 Then
                                tmpCLA(tmpCountPerLine, i).Ansi = -1
                                tmpCLA(tmpCountPerLine, i).Color = tmpCL.Color
                                tmpCLA(tmpCountPerLine, i).BColor = tmpCL.BColor
                            End If
                            tmpCountPerLine = tmpCountPerLine + tmpbyte
                            'ÅýCPU³Ý®§¤@¤U
                            DoEvents
                        Next L
                    End If
        Next j
    Next i
Exit Sub
out:
    Debug.Print "AnalyzeCC Error Out : " & Err.Description
End Sub



Public Function GetFC(ByVal inFCC As Byte) As Byte
    'FCC (FORE COLOR CODE)
    '©w¸q1;37 => 137
    Dim tmpbyte  As Byte
    
    Select Case inFCC Mod 100
        Case Is = 30
            tmpbyte = 0
        Case Is = 31
            tmpbyte = 4
        Case Is = 32
            tmpbyte = 2
        Case Is = 33
            tmpbyte = 6
        Case Is = 34
            tmpbyte = 1
        Case Is = 35
            tmpbyte = 5
        Case Is = 36
            tmpbyte = 3
        Case Is = 37
            tmpbyte = 7
        Case Else
            Debug.Print "SetFC µLªk¿ëÃÑ¨Ó·½"
    End Select
    '³B²z«G¦â
    If inFCC > 100 Then
        tmpbyte = tmpbyte + 8
    End If
    '¦^¶Ç
    GetFC = tmpbyte
    
End Function
Public Function GetBC(ByVal inBCC As Byte) As Byte
    'BCC (BACK COLOR CODE)

    Dim tmpbyte  As Byte
    
    Select Case inBCC
        Case Is = 40
            tmpbyte = 0
        Case Is = 41
            tmpbyte = 4
        Case Is = 42
            tmpbyte = 2
        Case Is = 43
            tmpbyte = 6
        Case Is = 44
            tmpbyte = 1
        Case Is = 45
            tmpbyte = 5
        Case Is = 46
            tmpbyte = 3
        Case Is = 47
            tmpbyte = 7
        Case Else
            Debug.Print "GETBC µLªk¿ëÃÑ¨Ó·½"
    End Select

    '¦^¶Ç
    GetBC = tmpbyte
    
End Function

Public Sub Post2Arrf(ByVal X As Integer, ByVal Y As Integer, Optional ByVal deBack As Byte)
    '¿ë§O¬O§_­n¥h­I
    If deBack = 1 Then
        Call CLArrayPaste_C(tmpCLA, Arrf, X, Y, OFP.CurrentPage)
    Else
        Call ObjLibPo(tmpCLA, Arrf, X, Y)
    End If
End Sub

Public Function AnalyzeCC_ByteArray(ByVal X As Integer, ByVal Y As Integer, ByVal maxX As Integer, ByVal maxY As Integer, Optional ByVal toClpB As Byte)
    'Åª¨ú°Å¶KÃ¯ªº°}¦C¶i¦æ¤ÀªR
    Dim CCMode As Byte  '0 ¤@¯ë
    Dim tmpStr As String '¦r¦ê¼È¦s¾¹
    Dim tmpInt As Integer
    Dim i As Long
    Dim Pointer_X As Integer, Pointer_Y As Integer, cutX As Integer, cutY As Integer
    Dim tmpbiByte(1) As Byte, biByteFlag As Byte, preFC_Byte As Byte, preBC_Byte As Byte, FC_Byte As Byte, BC_Byte As Byte, rawFClight_Byte As Byte, rawFC_Byte As Byte, rawBC_Byte As Byte
    Dim colorResetFlag As Byte
    Dim tmpint2 As Integer
    ReDim Ba_CLA(maxX, maxY)
    FC_Byte = 7
    BC_Byte = 0
    For i = 0 To UBound(ByteArray)
        If CCMode = 0 Then
            If ByteArray(i) = 0 Then
                Exit For
            ElseIf ByteArray(i) = 27 Then   '±m¦â½X»{ÃÒ
                If ByteArray(i + 1) = 91 Then   '½T©w±m¦â½X¿ë»{¦r¤¸«á¬°"["´N¶i¤J¼Ò¦¡1
                    i = i + 1
                    CCMode = 1
                End If
            ElseIf ByteArray(i) = 10 Then   '\n ´«¦æ
                If Pointer_Y < maxY Then
                    Pointer_X = 0
                    Pointer_Y = Pointer_Y + 1
                End If
                biByteFlag = 0
            ElseIf biByteFlag = 1 Then
                If Pointer_X < maxX Then
                    tmpbiByte(0) = ByteArray(i)
                    Call CopyMemory(tmpInt, tmpbiByte(0), 2)
                    Ba_CLA(Pointer_X, Pointer_Y).Ansi = tmpInt
                    Ba_CLA(Pointer_X, Pointer_Y).Color = preFC_Byte
                    Ba_CLA(Pointer_X, Pointer_Y).BColor = preBC_Byte
                    Ba_CLA(Pointer_X + 1, Pointer_Y).Ansi = -1
                    Ba_CLA(Pointer_X + 1, Pointer_Y).Color = FC_Byte
                    Ba_CLA(Pointer_X + 1, Pointer_Y).BColor = BC_Byte
                    If cutX < Pointer_X + 1 Then cutX = Pointer_X + 1 '¨ú³Ì¤jªºx
                    Pointer_X = Pointer_X + 2
                End If
                biByteFlag = 0
            ElseIf (ByteArray(i) >= 32 And ByteArray(i) <= 127) Then
                If Pointer_X < maxX Then
                    Ba_CLA(Pointer_X, Pointer_Y).Ansi = ByteArray(i)
                    Ba_CLA(Pointer_X, Pointer_Y).Color = FC_Byte
                    Ba_CLA(Pointer_X, Pointer_Y).BColor = BC_Byte
                    If cutX < Pointer_X Then cutX = Pointer_X   '¨ú³Ì¤jªºx
                    Pointer_X = Pointer_X + 1
                End If
            

            ElseIf ByteArray(i) >= 128 Then 'Âù¦ì¤¸¦r
                If Pointer_X < maxX Then
                    tmpbiByte(1) = ByteArray(i)
                    preFC_Byte = FC_Byte
                    preBC_Byte = BC_Byte
                    biByteFlag = 1
                End If

            End If
            
        
        ElseIf CCMode = 1 Then
            'Â^¨ú±m¦â½X
            If ByteArray(i) >= 48 And ByteArray(i) <= 57 Then   '0~9
                tmpStr = tmpStr & Chr(ByteArray(i))
            ElseIf ByteArray(i) = 59 Then   ';
                tmpint2 = Val(tmpStr)
                tmpStr = ""
                If tmpint2 = 0 Then
                    colorResetFlag = 1
                ElseIf tmpint2 = 1 Then
                    rawFClight_Byte = 1
                ElseIf tmpint2 >= 40 And tmpint2 <= 48 Then
                    rawBC_Byte = tmpint2
                ElseIf tmpint2 >= 30 And tmpint2 <= 38 Then
                    rawFC_Byte = tmpint2
                End If
            ElseIf ByteArray(i) = 77 Or ByteArray(i) = 109 Then 'M or m
                tmpint2 = Val(tmpStr)
                tmpStr = ""
                If tmpint2 = 0 Then
                    colorResetFlag = 1
                ElseIf tmpint2 = 1 Then
                    rawFClight_Byte = 1
                ElseIf tmpint2 >= 40 And tmpint2 <= 48 Then
                    rawBC_Byte = tmpint2
                ElseIf tmpint2 >= 30 And tmpint2 <= 38 Then
                    rawFC_Byte = tmpint2
                End If
                If colorResetFlag = 1 Then  'ªì©l¤Æ¦â±m
                    FC_Byte = 7
                    BC_Byte = 0
                    colorResetFlag = 0  'ÁÙ­ì¦¹¼Ð°O
                End If
                If rawFC_Byte = 0 Then  '³B²z«e´º¦â±m
                    If rawFClight_Byte <> 0 Then
                        FC_Byte = (FC_Byte Mod 8) + 8
                        rawFClight_Byte = 0  'ÁÙ­ì
                    End If
                Else
                    If rawFClight_Byte = 0 Then
                        If (FC_Byte \ 8) = 1 Then   '§PÂ_¦³µLÄ~©Ó«G¦â
                            FC_Byte = GetFC(rawFC_Byte) + 8
                        Else
                            FC_Byte = GetFC(rawFC_Byte)
                        End If
                    Else
                        FC_Byte = GetFC(rawFC_Byte) + 8
                        rawFClight_Byte = 0 'ÁÙ­ì
                    End If
                    rawFC_Byte = 0  'ÁÙ­ì
                End If
                If rawBC_Byte <> 0 Then
                    BC_Byte = GetBC(rawBC_Byte)
                    rawBC_Byte = 0
                End If
                CCMode = 0
            End If
            
        'ElseIf CCMode = 2 Then
            
        End If

    Next i
    'ReDim Preserve Ba_CLA(0 To cutX, 0 To Pointer_Y) As ColorLayer
    If toClpB = 1 Then
        ReDim CPArr(cutX, Pointer_Y)
        For j = 0 To Pointer_Y
            For i = 0 To cutX
                CPArr(i, j) = Ba_CLA(i, j)
            Next i
        Next j
    Else
        Call ObjLibPo_Area(Ba_CLA, Arrf, X, Y, cutX, Pointer_Y)
    End If
    ReDim Ba_CLA(0)
End Function
