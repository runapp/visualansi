Attribute VB_Name = "VAFile"
Public Type VAFileInfo
    IDC As Integer
    'IDC �ѧO�X
    filetype As Integer
    '�ɮ�����
    '1 �歶�e
    '2 �h���e
    '3 �ʵe
    '4 ����e
    '5 ����ʵe
    ArrXUbound As Integer
    ArrYUbound As Integer
    ArrZLenth As Integer
    '�ɮ׸�T�����c�Ƹ��
End Type


Public Sub VA_SaveFile(ByVal filepath As String, ByRef filedata() As ColorLayer, ByRef Fp As OpenedFilePropety, ByRef intimeLine() As Single)
On Error GoTo out
    Dim VAFI As VAFileInfo
    Dim OpFile As Integer
    '�g�J�ɮ׸�T
    VAFI.IDC = 828
    VAFI.filetype = Fp.filetype
    VAFI.ArrXUbound = UBound(filedata, 1)
    VAFI.ArrYUbound = UBound(filedata, 2)
    If VAFI.filetype = 2 Or VAFI.filetype = 3 Then
        VAFI.ArrZLenth = UBound(filedata, 3)
    End If
    
    'Debug.Print "VAFI.ArrXUbound =" & UBound(filedata, 1)
    'Debug.Print "VAFI.ArrYUbound =" & UBound(filedata, 2)
    OpFile = 1
    
    If Form1.FileSys.FileExists(filepath) = True Then
        '�Hbinary�}���ɮ� �n�屼��Ӫ� �~���|�d�U���e����T
       
        Kill filepath
        
        '�Y�ɮפ��s�b kill�|�o�Ϳ��~
    End If
        
    Open filepath For Binary As #OpFile
    Put #OpFile, 1, VAFI
    Put #OpFile, 11, filedata
    Put #OpFile, , intimeLine
    
    Debug.Print "savefile  " & filepath
    Close #OpFile
    Exit Sub
out:

End Sub

Public Sub VA_ReadFile(ByVal filepath As String, ByRef filedata() As ColorLayer, ByRef intimeLine() As Single)
    
    Dim VAFI As VAFileInfo
    Dim OpFile As Integer
    Dim tmpLen As Long
    
    OpFile = 1
    
    Open filepath For Binary As #OpFile
    Get #OpFile, 1, VAFI
    'ReDim filedata(VAFI.ArrXUbound, VAFI.ArrYUbound) As ColorLayer
    Select Case VAFI.filetype
        Case Is = 1
            Call Form1.SetSize(VAFI.ArrXUbound + 1, VAFI.ArrYUbound + 1, 1, 1)
        Case Is = 2
            Call Form1.SetSize(VAFI.ArrXUbound + 1, VAFI.ArrYUbound + 1, VAFI.ArrZLenth, 2)
        Case Is = 3
            Call Form1.SetSize(VAFI.ArrXUbound + 1, VAFI.ArrYUbound + 1, VAFI.ArrZLenth, 3)
    End Select
    
    Get #OpFile, 11, filedata
    tmpLen = 4 * (UBound(filedata, 1) + 1) * (UBound(filedata, 2) + 1) * UBound(filedata, 3) + 11 '�p��e����ƪ�����
    If tmpLen < LOF(OpFile) Then
        Debug.Print "Ū���ɶ��b" & tmpLen & "<" & LOF(OpFile)
        Get #OpFile, tmpLen, timeLine
    Else
        'Get #OpFile, tmpLen, timeLine
        Debug.Print "�S�ɶ��b" & tmpLen & "<" & LOF(OpFile)
    End If
    
    
    Close #OpFile


End Sub

Public Sub OpenNewFile()
Form4.Show vbModal
End Sub

Public Sub CreatNewFile(ByRef filedata() As ColorLayer, VAFI As VAFileInfo)
Debug.Print "�إߤ@�ӷs�ɮ� �ɮ׮榡" & VAFI.filetype & " �e: " & VAFI.ArrXUbound & " ��: " & VAFI.ArrYUbound
Select Case VAFI.filetype

    Case Is = 1
        Call Form1.SetSize(VAFI.ArrXUbound + 1, VAFI.ArrYUbound + 1, 1, 1)
    Case Is = 2
        Call Form1.SetSize(VAFI.ArrXUbound + 1, VAFI.ArrYUbound + 1, VAFI.ArrZLenth, 2)
    Case Is = 3
        Call Form1.SetSize(VAFI.ArrXUbound + 1, VAFI.ArrYUbound + 1, VAFI.ArrZLenth, 3)
End Select

End Sub


