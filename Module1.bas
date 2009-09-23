Attribute VB_Name = "Module1"
'�w�q�ϵe�򥻤���
Type ColorLayer
    BColor As Byte
    Color As Byte
    Ansi As Integer
End Type

'�w�q�ثe�ɮ��ݩ��ܼƼҲ�
Type OpenedFilePropety
    FilePath As String
    Closed As Boolean
    filetype As Integer
    CurrentPage As Integer
    IsChanged As Boolean
End Type

Type BiAnsi
    First As Byte
    Second As Byte
End Type
Type IntS
    v As String * 1
End Type
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)

Public Type POINTAPI
        x As Long
        y As Long
End Type
'�w�q�C��洫�\�઺�ܼƼҲ�
Type ExChColorType
    CurrentSel As Byte '�C��洫�\��_�ثe����n��⪺�C����
    Color(0 To 3) As Byte

End Type

'�w�q�e���ӷ����峹���ܼƼҲ�
Type ForStringType
    str As String   '�x�s�ثe���r��}�C
    StrLen(0 To 1) As Byte '�x�s�r�ꪺ���P�e
End Type

'�w�q�{���������ܼƼҲդ�
Public Type SysEnv
        PPA As Single  'pix per ansi
        Frontsize As Byte
        CheckSave As Byte
        HideSelect As Byte
        SSAnsi As Integer
        LastAnsi As Integer
        EdMode As Byte      '�ҿ��ø�Ϥu�㪺�Ҧ�
        ForeSource As Byte
        ForColor As Byte    '�ثe�ҿ�����e����
        BacColor As Byte    '�ثe�ҿ�����I����
        ExchC_Current As Byte
        ExChColor As ExChColorType  '�C��洫�\�઺�ܼƼҲ�
        ToolPBoxDown As Byte '�N�u��c�m��
        
        cDrawPos_X As Integer   '�ثeø�ϵ����ɪ���Ц�m
        cDrawPos_Y As Integer
        InputFlag As Byte
        CopyingFlag As Byte '��ܥ��b�ƻs
        
        '�즲����
        Move_Draging As Boolean
        MD_X As Integer '�즲��m���
        MD_Y As Integer
        
End Type
'������mAPI
Public Const HWND_TOPMOST = -1
Public Const HWND_TOP = 0
Public Const HWND_NOTOPMOST = -2
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_SHOWWINDOW = &H40
Public Declare Sub SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)

Public Declare Function Polygon Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long

Public Arrf() As ColorLayer '�w�q�D�n��ưO����
Public timeLine() As Single '�w�q�ɶ��b
Public CPArr() As ColorLayer   '�ƻs�K�W�ҥΤ��O����

Public OFP As OpenedFilePropety '�إߥثe�ɮ��ݩ��ܼƼҲ�
Public SysInfo As SysEnv '�إߵ{�������ܼƼҲ�
Public FString As ForStringType '�إ߫e�����峹���ܼ�(�]���פ��w�|�v�T��]�w�� �G�W�ߩ󤽥������ܼ�)
Public Sub Main()
'Form9.Show
Form8.Show
Dim PauseTime, Start, Finish, TotalTime
    PauseTime = 1 / 86400   ' �]�w�Ȱ��ɶ�
    Start = Now    ' �]�w�}�l�Ȱ����ɨ�C
    Do While Now < Start + PauseTime
        DoEvents    ' �N�{�������v�����䥦�{���C
    Loop
    If Command <> "" Then
        If Left(Command, 1) = "*" Then  '����Ұʤ���Ͻs�边���R�O
            Load Form14
        Else
            'Form1.MDIChild = True
            Load Form1
            'Form1.Show
        End If
    Else
        Load Form1
    End If
'Form7.Show
End Sub

Public Function ReadConfic(ByVal FilePosition As String) As SysEnv
    Dim ConficDS As SysEnv
    Dim ConficFile As Integer
    ConficFile = 3
    Open FilePosition For Binary As #ConficFile
    
    Get #ConficFile, 1, ConficDS
    Close ConficFile
    ReadConfic = ConficDS
End Function
Public Sub WriteConfic(ByVal FilePosition As String, ByRef ConficDS As SysEnv)
    Dim ConficFile As Integer
    ConficFile = 3
    Open FilePosition For Binary As #ConficFile
    
    Put #ConficFile, 1, ConficDS
    Close ConficFile
End Sub
Public Function Tlen(ByVal s As String) As Byte
    Tlen = LenB(StrConv(s, vbFromUnicode))
End Function
Public Function QBCToAnsiC(ByVal QBC As Integer) As String
    Dim AnsiC As String
    
    Select Case QBC Mod 8
        Case Is = 0
            AnsiC = "30"
        Case Is = 1
            AnsiC = "34"
        Case Is = 2
            AnsiC = "32"
        Case Is = 3
            AnsiC = "36"
        Case Is = 4
            AnsiC = "31"
        Case Is = 5
            AnsiC = "35"
        Case Is = 6
            AnsiC = "33"
        Case Is = 7
            AnsiC = "37"
    End Select
    
    If QBC \ 8 = 1 Then
        AnsiC = "1;" & AnsiC
    Else
        AnsiC = ";" & AnsiC
    End If
    'If QBC = 15 Then AnsiC = "1"
    QBCToAnsiC = AnsiC
End Function
Public Function QBCToAnsiBC(ByVal QBC As Integer) As String
    Dim AnsiC As String
    
    Select Case QBC Mod 8
        Case Is = 0
            AnsiC = "40"
        Case Is = 1
            AnsiC = "44"
        Case Is = 2
            AnsiC = "42"
        Case Is = 3
            AnsiC = "46"
        Case Is = 4
            AnsiC = "41"
        Case Is = 5
            AnsiC = "45"
        Case Is = 6
            AnsiC = "43"
        Case Is = 7
            AnsiC = "47"
    End Select
    
    QBCToAnsiBC = AnsiC
End Function

