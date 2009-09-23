Attribute VB_Name = "ByteArrayMudule"
Public ByteArray() As Byte
Private BA_EnlargeSize As Long
Private BA_Pointer As Long
Private BA_Ubound As Long

Public Sub BA_SetDefault()
    BA_EnlargeSize = 1048576
    BA_Pointer = 0
End Sub

Public Sub BA_Reset()
    ReDim ByteArray(BA_EnlargeSize - 1)
    BA_Pointer = 0
End Sub

Public Sub BA_Enlarge()
     ReDim Preserve ByteArray(BA_Ubound + BA_EnlargeSize)
     BA_Ubound = BA_Ubound + BA_EnlargeSize
End Sub

Public Sub BA_Put(ByVal tByte As Byte)
    ByteArray(BA_Pointer) = tByte
    BA_Pointer = BA_Pointer + 1
    If tByte = 0 Then
        Debug.Print "發現0"
    End If
    If BA_Pointer > BA_EnlargeSize Then
        Call BA_Enlarge
    End If
     
End Sub
Public Sub BA_Put_Str(ByVal Bstr As String)
    Dim tmpBArr() As Byte
    Dim tmpBlen As Integer
    Dim i As Integer
    If Bstr = "" Then Exit Sub
    tmpBArr = StrConv(Bstr, vbFromUnicode)
    tmpBlen = UBound(tmpBArr)
    If BA_Pointer + tmpBlen + 2 > BA_EnlargeSize Then
        Call BA_Enlarge
    End If
    For i = 0 To tmpBlen
        ByteArray(BA_Pointer) = tmpBArr(i)
        BA_Pointer = BA_Pointer + 1
        If tmpBArr(i) = 0 Then
            Debug.Print "發現0"
        End If
       
    Next i
    'Debug.Print "OK"
End Sub
Public Function BA_CutTail() As Long
    '實際資料到BA_Pointer-1但是因為ANSI字串結尾要是char(0)
    ReDim Preserve ByteArray(BA_Pointer)
    BA_Ubound = BA_Pointer
    BA_CutTail = BA_Pointer + 1 '回傳陣列大小
End Function


Public Sub BA_View()
    Debug.Print "BA_View"
End Sub
