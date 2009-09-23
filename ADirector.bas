Attribute VB_Name = "ADirector"


Public Sub DelPage(ByVal page As Integer)
    Dim oriX As Integer
    Dim oriY As Integer
    Dim oriZ As Integer
    oriX = UBound(Arrf, 1)
    oriY = UBound(Arrf, 2)
    oriZ = UBound(Arrf, 3)
    If page <> oriZ Then
        For i = page To oriZ - 1
            
            For j = 0 To oriX
                For k = 0 To oriY
                    Arrf(j, k, i) = Arrf(j, k, i + 1)
                    DoEvents
                Next k
            Next j
            timeLine(i) = timeLine(i + 1)
        Next i
    End If
    ReDim Preserve Arrf(oriX, oriY, 1 To oriZ - 1)
    ReDim Preserve timeLine(1 To oriZ - 1)
End Sub

Public Sub ExChPage(ByVal page1 As Integer, ByVal page2 As Integer)
    Dim tmpC As ColorLayer
    Dim oriX As Integer
    Dim oriY As Integer
    'Dim oriZ As Integer
    oriX = UBound(Arrf, 1)
    oriY = UBound(Arrf, 2)
    'oriZ = UBound(Arrf, 3)
        For j = 0 To oriX
            For k = 0 To oriY
                tmpC = Arrf(j, k, page1)
                Arrf(j, k, page1) = Arrf(j, k, page2)
                Arrf(j, k, page2) = tmpC
                DoEvents
            Next k
        Next j
End Sub

Public Sub InsertBlank(ByVal page As Integer)
    Dim tmpC As ColorLayer
    Dim oriX As Integer
    Dim oriY As Integer
    Dim oriZ As Integer
    oriX = UBound(Arrf, 1)
    oriY = UBound(Arrf, 2)
    oriZ = UBound(Arrf, 3)
    ReDim Preserve Arrf(oriX, oriY, 1 To oriZ + 1)
    ReDim Preserve timeLine(1 To oriZ + 1)
    tmpC.Color = 7
    For i = oriZ To page Step -1
        
        For j = 0 To oriX
            For k = 0 To oriY
                Arrf(j, k, i + 1) = Arrf(j, k, i)
                DoEvents
            Next k
        Next j
        timeLine(i + 1) = timeLine(i)
    Next i
        
        For j = 0 To oriX
            For k = 0 To oriY
                Arrf(j, k, page) = tmpC
                DoEvents
            Next k
        Next j
    
End Sub
