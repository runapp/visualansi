Attribute VB_Name = "ObjectLib"
Public ObjCA() As ColorLayer
Type ObjListDataType
    filepath As String * 100
    ObjName As String * 30
End Type
Public ObjList() As ObjListDataType
Public Sub CopyToObjLib(SourceArray() As ColorLayer, ByVal ObjName As String, ByVal X1 As Integer, ByVal Y1 As Integer, ByVal X2 As Integer, ByVal Y2 As Integer)
Dim AddLen As Boolean
AddLen = False
For i = Y1 To Y2
    If SourceArray(X2 + 1, i, OFP.CurrentPage).Ansi = -1 Then AddLen = True

Next i
Debug.Print "X2 =" & X2
If AddLen Then X2 = X2 + 1
Debug.Print "Addlen=" & AddLen
Debug.Print "X2 =" & X2

ReDim ObjCA(X2 - X1, Y2 - Y1) As ColorLayer
Call ObjCAfPreValue

For i = 0 To Y2 - Y1

    For j = 0 To X2 - X1
        ObjCA(j, i) = SourceArray(j + X1, i + Y1, OFP.CurrentPage)
        Debug.Print "ObjCA(" & j & "," & i & ").Ansi=" & ObjCA(j, i).Ansi
        DoEvents
    Next j

    
Next i
Dim Fp As OpenedFilePropety
Fp.filetype = 1
Call VA_SaveFile(App.Path & "\" & ObjName & ".vaf", ObjCA, Fp, timeLine)
'ReDim ObjCA(0)

End Sub

Public Sub ObjLibPo(ObjArray() As ColorLayer, TragetArray() As ColorLayer, ByVal X As Integer, ByVal Y As Integer)
    'ObjArray為2D
    'TragetArray為3D
    On Error GoTo out
    Debug.Print "UBound(ObjArray, 2)=" & UBound(ObjArray, 2)
    Debug.Print "UBound(ObjArray, 1)=" & UBound(ObjArray, 1)
    Dim dX As Integer, dY As Integer
    dX = UBound(ObjArray, 1)
    dY = UBound(ObjArray, 2)
    Dim maxX As Integer
    maxX = UBound(TragetArray, 1)
    'Call Form1.DoErease_Area_In(X, Y, X + dX, Y + dY)
    For j = Y To Y + dY
        For i = X To X + dX
            Call DoErease_A(Arrf, i, j, OFP.CurrentPage)
        Next i
    Next j
    'DoErease_A
    For i = 0 To dY
        For j = 0 To dX
            'If ObjArray(j, i).Ansi <> -1 Then
                'Call DoDraw_A(Arrf, X + j, Y + i, OFP.CurrentPage, Chr(ObjArray(j, i).Ansi), ObjArray(j, i).Color)
                'Call DoDrawBC_A(Arrf, X + j, Y + i, OFP.CurrentPage, ObjArray(j, i).BColor)
                If X + j = maxX And j <> dX Then
                    If ObjArray(j + 1, i).Ansi <> -1 Then
                        TragetArray(X + j, Y + i, OFP.CurrentPage) = ObjArray(j, i)
                    End If
                ElseIf X + j < maxX Then

                    TragetArray(X + j, Y + i, OFP.CurrentPage) = ObjArray(j, i)
                End If
            'End If
            DoEvents
        Next j
        
    Next i
    Exit Sub
out:
    Debug.Print "ObjLibPo Error Out"
    Debug.Print "(j,i)=" & j & "," & i
    Resume Next
End Sub
Public Sub ObjLibPo_Area(ObjArray() As ColorLayer, TragetArray() As ColorLayer, ByVal X As Integer, ByVal Y As Integer, ByVal dX As Integer, ByVal dY As Integer)
    'ObjArray為2D
    'TragetArray為3D
    On Error GoTo out
    'Debug.Print "UBound(ObjArray, 2)=" & UBound(ObjArray, 2)
    'Debug.Print "UBound(ObjArray, 1)=" & UBound(ObjArray, 1)
    
    'dX = UBound(ObjArray, 1)
    'dY = UBound(ObjArray, 2)
    
    Call Form1.DoErease_Area(X, Y, X + dX, Y + dY)
    For i = 0 To dY
        For j = 0 To dX

            Call DoDraw_A(Arrf, X + j, Y + i, OFP.CurrentPage, Chr(ObjArray(j, i).Ansi), ObjArray(j, i).Color)
            Call DoDrawBC_A(Arrf, X + j, Y + i, OFP.CurrentPage, ObjArray(j, i).BColor)

            DoEvents
        Next j
        
    Next i
    Exit Sub
out:
    Debug.Print "ObjLibPo Error Out"
    Debug.Print "(j,i)=" & j & "," & i
    Resume Next
End Sub

Public Sub ObjList_Add(fileinfo As ObjListDataType)
    ReDim Preserve ObjList(Form1.List1.ListCount)
    ObjList(Form1.List1.ListCount) = fileinfo
    Form1.List1.AddItem fileinfo.ObjName
    
End Sub

Public Sub ObjList_Del(ByVal Index As Integer)

If Index = -1 Then Exit Sub
Dim tempobjlist() As ObjListDataType
ReDim tempobjlist(UBound(ObjList) - 1) As ObjListDataType
For i = 0 To Index - 1
    tempobjlist(i) = ObjList(i)
Next i
For i = Index To UBound(ObjList) - 1
    tempobjlist(i) = ObjList(i + 1)
Next i
ReDim ObjList(UBound(ObjList) - 1) As ObjListDataType
ObjList = tempobjlist
Form1.List1.RemoveItem Index
If Form1.List1.ListCount = Index Then Index = Index - 1

Form1.List1.ListIndex = Index


End Sub



Public Sub Obj_ReadFile(ByVal filepath As String, ByRef filedata() As ColorLayer)
Dim VAFI As VAFileInfo
Dim OpFile As Integer

OpFile = 30

Open filepath For Binary As #OpFile
Get #OpFile, 1, VAFI
ReDim filedata(VAFI.ArrXUbound, VAFI.ArrYUbound) As ColorLayer
Debug.Print "VAFI.ArrXUbound=" & VAFI.ArrXUbound
Debug.Print "VAFI.ArrYUbound=" & VAFI.ArrYUbound
Get #OpFile, 11, filedata

Close #OpFile


End Sub

Public Sub ObjList_Read()
On Error GoTo out
Dim ObjListFile As Integer
Dim tempInt As Integer
ObjListFile = 3
Open App.Path & "\" & "ObjectList.dat" For Binary As #ObjListFile
    Get #ObjListFile, 1, tempInt
        Debug.Print "ObjList Count (Read)" & tempInt
    ReDim ObjList(tempInt) As ObjListDataType
    Get #ObjListFile, 11, ObjList

Close #ObjListFile
For i = 0 To tempInt
    Form1.List1.AddItem ObjList(i).ObjName
    DoEvents

Next i
Exit Sub
out:
Debug.Print "ObjList_Read Error Out"
End Sub
Public Sub ObjList_Save()
Dim ObjListFile As Integer
Dim tempInt As Integer
ObjListFile = 3
Open App.Path & "\" & "ObjectList.dat" For Binary As #ObjListFile
    tempInt = UBound(ObjList)
    Put #ObjListFile, 1, tempInt
    Debug.Print "ObjList Count (Save)" & tempInt
    Put #ObjListFile, 11, ObjList
Close #ObjListFile
End Sub
Public Sub ObjCAfPreValue()

    For j = 0 To UBound(ObjCA, 2)
        For i = 0 To UBound(ObjCA, 1)
            ObjCA(i, j).Color = 7
            DoEvents
        Next i
    Next j

End Sub
