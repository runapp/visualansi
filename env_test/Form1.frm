VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "���ҶE�_ for VA"
   ClientHeight    =   3375
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5775
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3375
   ScaleWidth      =   5775
   StartUpPosition =   2  '�ù�����
   Begin VB.CommandButton Command1 
      Caption         =   "�}�l����"
      Height          =   375
      Left            =   315
      TabIndex        =   1
      Top             =   525
      Width           =   1050
   End
   Begin VB.TextBox Text1 
      Height          =   2595
      Left            =   2085
      MultiLine       =   -1  'True
      ScrollBars      =   2  '�������b
      TabIndex        =   0
      Top             =   390
      Width           =   2880
   End
   Begin VB.Label Label2 
      Caption         =   "Nerv.Studio 2004-2005"
      Height          =   210
      Left            =   3900
      TabIndex        =   3
      Top             =   3165
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   555
      Left            =   210
      Picture         =   "Form1.frx":030A
      Stretch         =   -1  'True
      Top             =   2475
      Width           =   1680
   End
   Begin VB.Label Label1 
      Caption         =   "���i"
      Height          =   225
      Left            =   2145
      TabIndex        =   2
      Top             =   165
      Width           =   570
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type OSVERSIONINFO
        dwOSVersionInfoSize As Long
        dwMajorVersion As Long
        dwMinorVersion As Long
        dwBuildNumber As Long
        dwPlatformId As Long
        szCSDVersion As String * 128      '  Maintenance string for PSS usage
End Type

Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" _
        (lpVersionInformation As OSVERSIONINFO) As Long
Private Declare Function GetComputerName Lib "kernel32" Alias _
        "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Const VER_PLATFORM_WIN32_NT = 2
Private Const VER_PLATFORM_WIN32_WINDOWS = 1
Private Const VER_PLATFORM_WIN32s = 0
Private Declare Function GetSystemDirectory Lib "kernel32" Alias _
        "GetSystemDirectoryA" (ByVal lpBuffer As String, _
        ByVal nSize As Long) As Long
Dim EXEabletag As Byte


Private Sub GETOSINFO()
On Error Resume Next
Dim len5 As Long, aa As Long
Dim cmprName As String
Dim osver As OSVERSIONINFO

'���oComputer Name
cmprName = String(255, 0)
len5 = 256
aa = GetComputerName(cmprName, len5)
cmprName = Left(cmprName, InStr(1, cmprName, Chr(0)) - 1)
Text1.Text = "�q���W�� = " & cmprName

'���oOS������
osver.dwOSVersionInfoSize = Len(osver)
aa = GetVersionEx(osver)
Text1.Text = Text1.Text & vbCrLf & "MajorVersion " & osver.dwMajorVersion
Text1.Text = Text1.Text & vbCrLf & "MinorVersion " & osver.dwMinorVersion & vbCrLf & "�@�~�t�� "

Select Case osver.dwPlatformId
Case ER_PLATFORM_WIN32s
    Text1.Text = Text1.Text & "Microsoft Win32s "
Case VER_PLATFORM_WIN32_WINDOWS
    If (osver.dwMajorVersion = 4) And (osver.dwMinorVersion = 0) Then
         Text1.Text = Text1.Text & "Microsoft Windows 95 "
        If (Mid(osver.szCSDVersion, 2, 1) = "C") Then
            Text1.Text = Text1.Text & "OSR2 "
        End If
    ElseIf (osver.dwMajorVersion = 4) And (osver.dwMinorVersion = 10) Then
        If Mid(osver.szCSDVersion, 2, 1) = "A" Then
            Text1.Text = Text1.Text & "Microsoft Windows 98 SE"
        Else
            Text1.Text = Text1.Text & "Microsoft Windows 98"
        End If

    ElseIf (osver.dwMajorVersion = 4) And (osver.dwMinorVersion = 90) Then
        Text1.Text = Text1.Text & ("Microsoft Windows Me ")
    End If
Case VER_PLATFORM_WIN32_NT
    If osver.dwMajorVersion <= 4 Then
           Text1.Text = Text1.Text & "Microsoft Windows NT "

    ElseIf (osver.dwMajorVersion = 5) And (osver.dwMinorVersion) = 0 Then
            Text1.Text = Text1.Text & "Microsoft Windows 2000 "

    ElseIf (osver.dwMajorVersion = 5) And (osver.dwMinorVersion = 1) Then
            Text1.Text = Text1.Text & "Windows XP"
    End If
End Select
 Text1.Text = Text1.Text & vbCrLf & "===================" & vbCrLf
End Sub
Private Sub Command1_Click()
EXEabletag = 0
Call GETOSINFO
Call GETSYSDIR
Call MSFlexGridTest
Call CDialogTest
Call FileSysTest
Call CommonControlTest
Call RichTextTest

If EXEabletag = 0 Then
    MsgBox "�A���t�Τ��w�]�tVA0.9.7�H�U��������һ��ɮ�", 64, "���է���"
    Text1.Text = Text1.Text & "�ɥR:" & vbCrLf & "�p�G���M�L�k����VA,�нT�w�q�����a�ϳ]�m(����x->�a�Ͽﶵ->�@��)�� ����(�x�W) ,�����دS�O�P��superlubu(�l�L�f��)����U����" & vbCrLf & "===================" & vbCrLf
Else
    MsgBox "�o�{�ʤ�" & EXEabletag & "���ɮ�,�Шϥ�VA������w����", 64, "���է���"
End If
End Sub

Public Sub MSFlexGridTest()
On Error GoTo out
Load Form2

Unload Form2
Text1.Text = Text1.Text & "MSFlexGrid ���`���J" & vbCrLf & "===================" & vbCrLf
Exit Sub
out:
Text1.Text = Text1.Text & "MSFlexGrid ���J����" & vbCrLf & "���~�T��:" & Err.Description
EXEabletag = EXEabletag + 1

End Sub

Public Sub CDialogTest()
On Error GoTo out
Load Form3

Unload Form3
Text1.Text = Text1.Text & "CDialog ���`���J" & vbCrLf & "===================" & vbCrLf
Exit Sub
out:
Text1.Text = Text1.Text & "CDialog ���J����" & vbCrLf & "���~�T��:" & Err.Description
EXEabletag = EXEabletag + 1
End Sub
Public Sub FileSysTest()
On Error GoTo out
Dim filesys As New FileSystemObject
Text1.Text = Text1.Text & "FileSystemObject ���`���J" & vbCrLf & "===================" & vbCrLf
Exit Sub
out:
Text1.Text = Text1.Text & "FileSystemObject ���J����" & vbCrLf & "���~�T��:" & Err.Description
EXEabletag = EXEabletag + 1
End Sub

Public Sub GETSYSDIR()
Dim SysPath As String
SysPath = String(255, 0)
len5 = GetSystemDirectory(SysPath, 256)
SysPath = Left(SysPath, InStr(1, SysPath, Chr(0)) - 1)
Text1.Text = Text1.Text & "System Path : " & SysPath & vbCrLf & "===================" & vbCrLf

End Sub
Public Sub CommonControlTest()
On Error GoTo out
Load Form4

Unload Form4
Text1.Text = Text1.Text & "CommonControl ���`���J" & vbCrLf & "===================" & vbCrLf
Exit Sub
out:
Text1.Text = Text1.Text & "CommonControl ���J����" & vbCrLf & "���~�T��:" & Err.Description
EXEabletag = EXEabletag + 1
End Sub

Private Sub Image1_Click()
MsgBox "http://98.to/�s", 64, "�w����{"
End Sub

Public Sub RichTextTest()
On Error GoTo out
Load Form5

Unload Form5
Text1.Text = Text1.Text & "RichTextBox ���`���J" & vbCrLf & "===================" & vbCrLf
Exit Sub
out:
Text1.Text = Text1.Text & "RichTextBox ���J����" & vbCrLf & "���~�T��:" & Err.Description
EXEabletag = EXEabletag + 1
End Sub


