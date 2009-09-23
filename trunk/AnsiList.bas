Attribute VB_Name = "AnsiList"
Public Sub LoadAnsi(ByVal filepath As String, Msfg As MSFlexGrid)
On Error GoTo out
Msfg.Visible = False
Dim Apointer() As Byte
Dim tstr As String
AnsiFile = 1
Dim counter As Integer

Open filepath For Binary As #AnsiFile

ReDim Apointer(1 To LOF(AnsiFile)) As Byte
Get #AnsiFile, 1, Apointer


tstr = StrConv(Apointer, vbUnicode)

For i = 1 To Len(tstr)

    Call IntoGrid(Msfg, counter, Mid(tstr, i, 1))
    DoEvents

Next i


Close AnsiFile
Msfg.Visible = True
Exit Sub
out:
MsgBox "載入字元表時發生錯誤 請檢查Ansi.txt是否存在", 16, "錯誤"
End Sub
Public Sub LoadAnsiLs(ByVal filepath As String, LB As ListBox)
Dim Apointer() As Byte
Dim tstr As String
AnsiFile = 1
Dim counter As Integer

Open filepath For Binary As #AnsiFile

ReDim Apointer(1 To LOF(AnsiFile)) As Byte
Get #AnsiFile, 1, Apointer

tstr = StrConv(Apointer, vbUnicode)

For i = 1 To Len(tstr)

    LB.AddItem Mid(tstr, i, 1)
    DoEvents

Next i
Close AnsiFile
End Sub


Public Sub IntoGrid(Msfg As MSFlexGrid, ByRef counter As Integer, ByVal Ansi As String)

If Ansi = " " Then Exit Sub
If counter Mod 10 = 0 Then
    Msfg.AddItem ""

End If
Msfg.TextMatrix(counter \ 10, counter Mod 10) = Ansi
counter = counter + 1


End Sub

Public Sub ExChList(LB As ListBox, ByVal x1 As Integer, ByVal X2 As Integer)

If X2 = -1 Then Exit Sub
Dim tempstr As String

tempstr = LB.List(x1)
LB.List(x1) = LB.List(X2)
LB.List(X2) = tempstr


End Sub
Public Sub ExChGrid(Msfg As MSFlexGrid, ByVal Row1 As Integer, ByVal Col1 As Integer, ByVal Row2 As Integer, ByVal Col2 As Integer)
Dim tempansi1 As String * 1
Dim tempansi2 As String * 1
If Row2 < 0 Or Col2 < 0 Or Col2 >= 10 Then Exit Sub
tempansi1 = Msfg.TextMatrix(Row1, Col1)
tempansi2 = Msfg.TextMatrix(Row2, Col2)
If tempansi2 = "" Then Exit Sub
Msfg.TextMatrix(Row1, Col1) = tempansi2
Msfg.TextMatrix(Row2, Col2) = tempansi1



End Sub
