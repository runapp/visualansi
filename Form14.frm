VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form14 
   AutoRedraw      =   -1  'True
   Caption         =   "Visual Ansi �����"
   ClientHeight    =   5025
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   7470
   Icon            =   "Form14.frx":0000
   LinkTopic       =   "Form14"
   ScaleHeight     =   335
   ScaleMode       =   3  '����
   ScaleWidth      =   498
   StartUpPosition =   2  '�ù�����
   Begin VB.Frame Frame2 
      Caption         =   "�ĪG"
      Height          =   2670
      Index           =   3
      Left            =   5475
      TabIndex        =   14
      Top             =   915
      Visible         =   0   'False
      Width           =   1950
      Begin VB.CommandButton Command9 
         Caption         =   "�k��"
         Height          =   330
         Index           =   3
         Left            =   210
         TabIndex        =   27
         Top             =   1830
         Width           =   1530
      End
      Begin VB.CommandButton Command9 
         Caption         =   "�Ϧ�"
         Height          =   330
         Index           =   2
         Left            =   210
         TabIndex        =   26
         Top             =   1335
         Width           =   1530
      End
      Begin VB.CommandButton Command9 
         Caption         =   "����½��"
         Height          =   330
         Index           =   1
         Left            =   210
         TabIndex        =   25
         Top             =   840
         Width           =   1530
      End
      Begin VB.CommandButton Command9 
         Caption         =   "����½��"
         Height          =   330
         Index           =   0
         Left            =   210
         TabIndex        =   24
         Top             =   375
         Width           =   1530
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "�Ϥ�"
      Height          =   2670
      Index           =   5
      Left            =   5595
      TabIndex        =   32
      Top             =   1350
      Visible         =   0   'False
      Width           =   1950
      Begin VB.CommandButton Command12 
         Caption         =   "�ഫ�Ϥ���e��"
         Height          =   330
         Index           =   2
         Left            =   210
         TabIndex        =   36
         Top             =   1860
         Width           =   1530
      End
      Begin VB.CommandButton Command12 
         Caption         =   "�]���I��"
         Height          =   330
         Index           =   1
         Left            =   210
         TabIndex        =   35
         Top             =   1365
         Width           =   1530
      End
      Begin VB.CommandButton Command12 
         Caption         =   "�˵�/�վ�Ϥ�"
         Height          =   330
         Index           =   0
         Left            =   210
         TabIndex        =   34
         Top             =   855
         Width           =   1530
      End
      Begin VB.CommandButton Command11 
         Caption         =   "���J�Ϥ�"
         Height          =   330
         Left            =   210
         TabIndex        =   33
         Top             =   360
         Width           =   1530
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "�ƻs�K�W"
      Height          =   2670
      Index           =   4
      Left            =   5475
      TabIndex        =   15
      Top             =   1110
      Visible         =   0   'False
      Width           =   1950
      Begin VB.CheckBox Check1 
         Caption         =   "�h�I"
         Height          =   180
         Left            =   210
         TabIndex        =   31
         Top             =   2025
         Width           =   1230
      End
      Begin VB.CommandButton Command10 
         Caption         =   "�K�W"
         Height          =   330
         Index           =   2
         Left            =   210
         TabIndex        =   30
         Top             =   1485
         Width           =   1530
      End
      Begin VB.CommandButton Command10 
         Caption         =   "�ŤU"
         Height          =   330
         Index           =   1
         Left            =   210
         TabIndex        =   29
         Top             =   900
         Width           =   1530
      End
      Begin VB.CommandButton Command10 
         Caption         =   "�ƻs"
         Height          =   330
         Index           =   0
         Left            =   210
         TabIndex        =   28
         Top             =   375
         Width           =   1530
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "�����"
      Height          =   2670
      Index           =   1
      Left            =   5370
      TabIndex        =   9
      Top             =   675
      Visible         =   0   'False
      Width           =   1950
      Begin VB.CommandButton Command7 
         Caption         =   "�M���ҿ�ϰ�"
         Height          =   330
         Index           =   1
         Left            =   210
         TabIndex        =   23
         Top             =   615
         Width           =   1395
      End
      Begin VB.CommandButton Command7 
         Caption         =   "�M���㭶"
         Height          =   330
         Index           =   0
         Left            =   195
         TabIndex        =   22
         Top             =   1440
         Width           =   1395
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "����"
      Height          =   2670
      Index           =   2
      Left            =   5205
      TabIndex        =   13
      Top             =   480
      Visible         =   0   'False
      Width           =   1950
      Begin VB.CommandButton Command6 
         Caption         =   "���"
         Enabled         =   0   'False
         Height          =   285
         Left            =   1170
         TabIndex        =   21
         Top             =   735
         Width           =   645
      End
      Begin VB.OptionButton Option1 
         Caption         =   "����϶�"
         Height          =   180
         Index           =   1
         Left            =   120
         TabIndex        =   20
         Top             =   780
         Width           =   1065
      End
      Begin VB.OptionButton Option1 
         Caption         =   "�I��"
         Height          =   180
         Index           =   0
         Left            =   105
         TabIndex        =   19
         Top             =   435
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.Label Label5 
         BackStyle       =   0  '�z��
         Caption         =   "--->"
         Height          =   225
         Left            =   780
         TabIndex        =   18
         Top             =   1710
         Width           =   300
      End
      Begin VB.Label Label4 
         BackColor       =   &H00000000&
         Height          =   495
         Index           =   1
         Left            =   1140
         MouseIcon       =   "Form14.frx":030A
         MousePointer    =   99  '�ۭq
         TabIndex        =   17
         Top             =   1590
         Width           =   480
      End
      Begin VB.Label Label4 
         BackColor       =   &H00000000&
         BorderStyle     =   1  '��u�T�w
         Height          =   525
         Index           =   0
         Left            =   210
         MouseIcon       =   "Form14.frx":045C
         MousePointer    =   99  '�ۭq
         TabIndex        =   16
         Top             =   1560
         Width           =   540
      End
   End
   Begin MSComDlg.CommonDialog CDialog1 
      Left            =   3300
      Top             =   30
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      Caption         =   "�e��"
      Height          =   2670
      Index           =   0
      Left            =   5175
      TabIndex        =   8
      Top             =   270
      Width           =   1950
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   3270
      Left            =   720
      ScaleHeight     =   214
      ScaleMode       =   3  '����
      ScaleWidth      =   291
      TabIndex        =   5
      Top             =   810
      Width           =   4425
      Begin VB.Shape Shape3 
         BorderColor     =   &H00FF00FF&
         BorderStyle     =   3  '�I�u
         DrawMode        =   4  'Mask Not Pen
         Height          =   150
         Left            =   15
         Top             =   15
         Width           =   150
      End
   End
   Begin VB.Frame Frame4 
      Height          =   600
      Left            =   30
      TabIndex        =   3
      Top             =   -15
      Width           =   2595
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   390
         Left            =   60
         TabIndex        =   4
         Top             =   150
         Width           =   2400
         _ExtentX        =   4233
         _ExtentY        =   688
         ButtonWidth     =   609
         ButtonHeight    =   582
         ImageList       =   "ImgList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   6
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               ImageIndex      =   1
               Style           =   2
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               ImageIndex      =   2
               Style           =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               ImageIndex      =   3
               Style           =   2
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               ImageIndex      =   4
               Style           =   2
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               ImageIndex      =   5
               Style           =   2
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               ImageIndex      =   6
               Style           =   2
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame3 
      Height          =   705
      Left            =   45
      TabIndex        =   2
      Top             =   4170
      Width           =   5655
      Begin VB.ComboBox Combo1 
         Height          =   300
         Left            =   2265
         Style           =   2  '��¤U�Ԧ�
         TabIndex        =   37
         Top             =   240
         Width           =   870
      End
      Begin VB.CommandButton Command5 
         Caption         =   "���J�ťխ�"
         Height          =   285
         Left            =   4215
         TabIndex        =   12
         Top             =   240
         Width           =   1200
      End
      Begin VB.CommandButton Command2 
         Caption         =   "�U�@��"
         Height          =   285
         Index           =   1
         Left            =   3195
         TabIndex        =   11
         Top             =   255
         Width           =   840
      End
      Begin VB.CommandButton Command2 
         Caption         =   "�W�@��"
         Height          =   285
         Index           =   0
         Left            =   1395
         TabIndex        =   10
         Top             =   240
         Width           =   810
      End
      Begin VB.CommandButton Command3 
         Caption         =   "��s�e��"
         Height          =   300
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   1155
      End
   End
   Begin MSComctlLib.ImageList ImgList1 
      Left            =   2655
      Top             =   -15
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form14.frx":05AE
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form14.frx":094A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form14.frx":0CE6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form14.frx":19C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form14.frx":1CE6
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form14.frx":2002
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "��L"
      Height          =   2130
      Left            =   -15
      TabIndex        =   0
      Top             =   720
      Width           =   720
      Begin VB.PictureBox Pic1 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         DrawWidth       =   5
         FillStyle       =   4  '���W��k�U���׽u
         Height          =   1155
         Left            =   60
         MouseIcon       =   "Form14.frx":2CDE
         MousePointer    =   2  '�Q�r�Ϊ�
         ScaleHeight     =   4
         ScaleMode       =   0  '�ϥΪ̦ۭq
         ScaleWidth      =   1.6
         TabIndex        =   1
         Top             =   825
         Width           =   510
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  '���z��
         Height          =   360
         Left            =   120
         Top             =   300
         Width           =   390
      End
   End
   Begin VB.Label Label6 
      Height          =   195
      Left            =   1875
      TabIndex        =   38
      Top             =   600
      Width           =   2040
   End
   Begin VB.Label Label1 
      Height          =   195
      Left            =   870
      TabIndex        =   6
      Top             =   600
      Width           =   720
   End
   Begin VB.Menu Me_File 
      Caption         =   "�ɮ�(&F)"
      Begin VB.Menu Me_File_New 
         Caption         =   "�إ߷s��"
      End
      Begin VB.Menu Me_File_Open 
         Caption         =   "�}������"
      End
      Begin VB.Menu Me_File_Save 
         Caption         =   "�x�s�ɮ�"
      End
      Begin VB.Menu Me_File_Save_As 
         Caption         =   "�t�s�s��"
      End
   End
   Begin VB.Menu Me_Compile 
      Caption         =   "�sĶ(&C)"
   End
   Begin VB.Menu Me_About 
      Caption         =   "����(&A)"
   End
End
Attribute VB_Name = "Form14"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim hBrushColor(7) As Long
Public CDrawer As New CubeDrawerObject
Public SL As New SelectLine
Dim tmpPic As StdPicture
Dim ToolP As Frame
Dim cToolIndex As Integer
Dim ExChColorVar(2) As Byte '0����fromColor 1����toColor 2����ExChColorIndex
Dim cEdMode As Byte
Dim FixMouseMove(1) As Integer '0��x 1��y
Dim CPBflag As Byte '�����ƻs�K�W�Ҧ����A
Public IsChanged As Boolean

Public Sub SetColorBoard()
    '��ܽզ�L
    For i = 0 To 7
        Pic1.Line (i Mod 2, Fix(i / 2))-((i Mod 2) + 1, Fix(i / 2) + 1), QBColor(i), BF
    Next i

End Sub

Private Sub Command1_Click()
    Set tmpPic = LoadPicture(Text1.text)

End Sub

Private Sub Combo1_Click()

    Call CDrawer.GoToPage(Combo1.ListIndex + 1)
    Call Command2_Click(3)
End Sub

Private Sub Command10_Click(Index As Integer)
    Dim X1 As Integer, X2 As Integer, Y1 As Integer, Y2 As Integer
    Dim tmpInt As Integer
    X1 = SL.StartPoint_X
    Y1 = SL.StartPoint_Y
    X2 = SL.EndPoint_X
    Y2 = SL.EndPoint_Y
    If X1 > X2 Then
        tmpInt = X1
        X1 = X2
        X2 = tmpInt
    End If
    If Y1 > Y2 Then
        tmpInt = Y1
        Y1 = Y2
        Y2 = tmpInt
    End If
    Select Case Index
        Case 0
            Call CDrawer.CPB_Copy(X1, Y1, X2, Y2)
        
        Case 1
            Call CDrawer.CPB_Cut(X1, Y1, X2, Y2)
        
        Case 2
            If CPBflag = 1 Then
                Command10(2).Caption = "�K�W"
                Command10(0).Enabled = True
                Command10(1).Enabled = True
                CPBflag = 0
            Else
                Command10(2).Caption = "�����K�W"
                Command10(0).Enabled = False
                Command10(1).Enabled = False
                CPBflag = 1
            End If
    
    
    End Select
End Sub

Public Sub Command11_Click()
    On Error GoTo out
    CDialog1.DialogTitle = "���J�Ϥ�"
    CDialog1.Filter = "�䴩���Ϥ��榡(*.bmp;*.jpg;*.gif)|*.bmp;*.jpg;*.gif|�Ҧ��榡(*.*)|*.*"
    CDialog1.FileName = ""
    CDialog1.ShowOpen
    If Dir(CDialog1.FileName, vbHidden Or vbDirectory Or vbReadOnly Or vbSystem) = "" Then
        MsgBox "�ɮפ��s�b", 16, "���~"
    Else
        Call CDrawer.LoadIP(CDialog1.FileName)
        Form16.Show
        Call Form16.ReSize_UI
        Call Form16.ShowPic
    End If
    Exit Sub
out:
    
End Sub

Public Sub Command12_Click(Index As Integer)
    Select Case Index
        Case 0
            Form16.Show
        Case 1
            If CDrawer.HaveBG = 1 Then
                CDrawer.HaveBG = 0
                Command12(1).Caption = "�]���I��"
            Else
                CDrawer.HaveBG = 1
                Command12(1).Caption = "�����I��"
            End If
            Call CDrawer.ReShow(0, 0, CDrawer.W - 1, CDrawer.H - 1)
        Case 2
            '�Ϥ��ഫ��va�e���W
            Form17.Show
        'Case 4
            
            
    End Select
    
    
End Sub

Public Sub Command2_Click(Index As Integer)
'�e�@�� �U�@��
'�ǤJ��index��3��ܫ����椶������s
    If Index = 0 Then
        CDrawer.GoToPage (CDrawer.CurrentPage - 1)
        Combo1.ListIndex = CDrawer.CurrentPage - 1
    End If
    If Index = 1 Then
        CDrawer.GoToPage (CDrawer.CurrentPage + 1)
        Combo1.ListIndex = CDrawer.CurrentPage - 1
    End If
    If CDrawer.CurrentPage = CDrawer.Z Then
        Command2(1).Enabled = False
    Else
        Command2(1).Enabled = True
    End If
    If CDrawer.CurrentPage = 1 Then
        Command2(0).Enabled = False
    Else
        Command2(0).Enabled = True
    End If
    'Debug.Print CDrawer.CurrentPage - 1
    
  
End Sub

Private Sub Command3_Click()
    Call CDrawer.ReShow(0, 0, CDrawer.W - 1, CDrawer.H - 1)
End Sub





Private Sub Command5_Click()
    Call CDrawer.InsertPage(CDrawer.CurrentPage + 1)
    IsChanged = True
    Call SetMainCaption
    Call Refresh_ComboPage
    Call Command2_Click(1)  '�U�@��
End Sub

Private Sub Command6_Click()
    Dim X1 As Integer, X2 As Integer, Y1 As Integer, Y2 As Integer
    Dim tmpInt As Integer
    X1 = SL.StartPoint_X
    Y1 = SL.StartPoint_Y
    X2 = SL.EndPoint_X
    Y2 = SL.EndPoint_Y
    If X1 > X2 Then
        tmpInt = X1
        X1 = X2
        X2 = tmpInt
    End If
    If Y1 > Y2 Then
        tmpInt = Y1
        Y1 = Y2
        Y2 = tmpInt
    End If
    Call CDrawer.ExChColor_Area(X1, Y1, X2, Y2, ExChColorVar(0), ExChColorVar(1))
    IsChanged = True
    Call SetMainCaption
    'Debug.Print X1 & "," & Y1 & "," & X2 & "," & Y2
End Sub

Private Sub Command7_Click(Index As Integer)
    Dim X1 As Integer, X2 As Integer, Y1 As Integer, Y2 As Integer
    Dim tmpInt As Integer
    X1 = SL.StartPoint_X
    Y1 = SL.StartPoint_Y
    X2 = SL.EndPoint_X
    Y2 = SL.EndPoint_Y
    If X1 > X2 Then
        tmpInt = X1
        X1 = X2
        X2 = tmpInt
    End If
    If Y1 > Y2 Then
        tmpInt = Y1
        Y1 = Y2
        Y2 = tmpInt
    End If
    If Index = 1 Then
        Call CDrawer.Erease_Area(X1, Y1, X2, Y2)
    Else
        Call CDrawer.Erease_All
    End If
    IsChanged = True
    Call SetMainCaption
End Sub



Private Sub Command9_Click(Index As Integer)
    Dim X1 As Integer, X2 As Integer, Y1 As Integer, Y2 As Integer
    Dim tmpInt As Integer
    Call SL.CorrectPoints
    If SL.StartPoint_X < 0 Then SL.StartPoint_X = 0
    If SL.StartPoint_Y < 0 Then SL.StartPoint_Y = 0
    If SL.EndPoint_X >= CDrawer.W Then SL.EndPoint_X = CDrawer.W - 1
    If SL.EndPoint_Y >= CDrawer.H Then SL.EndPoint_Y = CDrawer.H - 1
    Call SL.DrawSelect
    X1 = SL.StartPoint_X
    Y1 = SL.StartPoint_Y
    X2 = SL.EndPoint_X
    Y2 = SL.EndPoint_Y
    'If x1 > X2 Then
    '    tmpInt = x1
    '    x1 = X2
    '    X2 = tmpInt
    'End If
    'If Y1 > Y2 Then
    '    tmpInt = Y1
    '    Y1 = Y2
    '    Y2 = tmpInt
    'End If
    Select Case Index
        Case 0
            Call CDrawer.Flip_H(X1, Y1, X2, Y2)
        Case 1
            Call CDrawer.Flip_V(X1, Y1, X2, Y2)
        Case 2
            Call CDrawer.FlipColor(X1, Y1, X2, Y2)
        Case 3
            
            Call CDrawer.Rotate_Right(X1, Y1, X2, Y2)
           
            SL.StartPoint_X = X1
            SL.StartPoint_Y = Y1
            SL.EndPoint_X = X1 + (Y2 - Y1)
            SL.EndPoint_Y = Y1 + (X2 - X1)
            'Call SL.CorrectPoints
            If SL.EndPoint_X >= CDrawer.W Then SL.EndPoint_X = CDrawer.W - 1
            If SL.EndPoint_Y >= CDrawer.H Then SL.EndPoint_Y = CDrawer.H - 1
            Call SL.DrawSelect
            Label6.Caption = "(" & SL.StartPoint_X & "," & SL.StartPoint_Y & ")-(" & SL.EndPoint_X & "," & SL.EndPoint_Y & ")"
    End Select
    IsChanged = True
    Call SetMainCaption
End Sub

Private Sub Form_Load()
    '�T���d�I
    
    prevWndProc = GetWindowLong(Me.hwnd, GWL_WNDPROC)
    SetWindowLong Me.hwnd, GWL_WNDPROC, AddressOf WndProc
    
    Call SetColorBoard
    CDrawer.TargetPB = Picture1
    SL.TragetShape = Shape3
    CDrawer.SetColor 7
    Set ToolP = Frame2(0)
    CDialog1.CancelError = True
    '�]�w�u���ݩʹw�]��
    cEdMode = 1
    Set ToolP = Frame2(0)
    cToolIndex = 0  '�ثe���ݩʤu�����(�Hframe��index�����X)
    Toolbar1.Buttons(cToolIndex + 1).Value = tbrPressed
    Call EnableEdit(False)
    Unload Form8
    Me.Show
    If Len(Command) <> 1 Then Call OpenFile_Command(Right(Command, Len(Command) - 1))
End Sub

Public Sub FormSetColor(ByVal X As Integer, ByVal Y As Integer)
    Dim tmpbyte As Byte
    tmpbyte = X + 2 * Y
    If cEdMode = 3 Then
        '�B�z����
        Label4(ExChColorVar(2)).BackColor = QBColor(tmpbyte)
        ExChColorVar(ExChColorVar(2)) = tmpbyte
    Else
        Call CDrawer.SetColor(tmpbyte)
        Shape2.BackColor = QBColor(tmpbyte)
    End If
End Sub

Private Sub Form_Paint()
    Debug.Print "Form_Paint"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call AskSave
    Unload Form16
    Unload Form6
    Unload Form17
    Unload Form18
    
End Sub

Private Sub Form_Resize()
    '���F�ѨM�@�ө_�Ǫ����D*�ҳ]�m
    '*���ܵ����j�p�� �Ϥ�������e���C��]�w�|�]��
    Call CDrawer.ReFreshColor
End Sub

Private Sub Form_Unload(Cancel As Integer)
    '�����T���d�I
        SetWindowLong Me.hwnd, GWL_WNDPROC, prevWndProc
End Sub

Private Sub Label2_Click()

End Sub

Private Sub Label4_Click(Index As Integer)
    ExChColorVar(2) = Index
    Label4(Index).BorderStyle = 1
    Label4(1 - Index).BorderStyle = 0
End Sub

Private Sub Me_About_Click()
    Form9.Show vbModal
End Sub

Private Sub Me_Compile_Click()
    If Form18.Visible = True Then
        Form18.Show
    Else
        Form18.Show vbModal
    End If
End Sub

Private Sub Me_File_New_Click()
    Call AskSave
    Form15.Show vbModal
End Sub

Private Sub Me_File_Open_Click()
On Error GoTo out
    Call AskSave
    CDialog1.DialogTitle = "�}������"
    CDialog1.Filter = "*.VAC(���e)|*.vac"
    CDialog1.FileName = ""
    CDialog1.ShowOpen
    If Dir(CDialog1.FileName, vbHidden Or vbDirectory Or vbReadOnly Or vbSystem) = "" Then
        MsgBox "�ɮפ��s�b", 16, "���~"
    Else
        Call CDrawer.OpenFile(CDialog1.FileName)
        Call Form16.ReSize_UI
        Call EnableEdit(True)
        Call Refresh_ComboPage
        Call Command2_Click(3)
        IsChanged = False
        Call SetMainCaption
    End If
    Call SetUIPos
    
    Exit Sub
out:
    Debug.Print "Me_File_Open_Click OUT "
End Sub


Private Sub Me_File_Save_Click()
On Error GoTo out
    If CDrawer.cFilepath = "" Then
        CDialog1.DialogTitle = "�x�s�ɮ�"
        CDialog1.Filter = "*.VAC(���e)|*.vac"
        CDialog1.FileName = ""
        CDialog1.ShowSave
        If Dir(CDialog1.FileName, vbHidden Or vbDirectory Or vbReadOnly Or vbSystem) <> "" Then
            If MsgBox("�ɮפw�g�s�b �аݧA�n�л\�L��?", 49, "�ɮפw�s�b") = vbCancel Then Exit Sub
            
        End If
        Call CDrawer.SaveFile(CDialog1.FileName)
    Else
        Call CDrawer.SaveFile(CDrawer.cFilepath)
    End If
    IsChanged = False
    Call SetMainCaption
    Call SetUIPos
    Exit Sub
out:
    Debug.Print "Me_File_Save_Click OUT "
End Sub

Private Sub Me_File_Save_As_Click()
On Error GoTo out
    CDialog1.DialogTitle = "�t�s�s��"
    CDialog1.Filter = "*.VAC(���e)|*.vac"
    CDialog1.FileName = ""
    CDialog1.ShowSave
    If Dir(CDialog1.FileName, vbHidden Or vbDirectory Or vbReadOnly Or vbSystem) <> "" Then
        If MsgBox("�ɮפw�g�s�b �аݧA�n�л\�L��?", 49, "�ɮפw�s�b") = vbCancel Then Exit Sub
    End If
    Call CDrawer.SaveFile(CDialog1.FileName)
    Exit Sub
out:


End Sub

Private Sub Option1_Click(Index As Integer)
    Command6.Enabled = Option1(1).Value
     
End Sub

Private Sub Pic1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo out
    intX = Fix(X)
    intY = Fix(Y)
    Call FormSetColor(intX, intY)
    
Exit Sub
out:
Debug.Print "Pic1_MouseDown Error Out : " & Err.Description
End Sub




Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo out
    intX = Fix(X)
    intY = Fix(Y)
    Dim tmpSLmode As Byte
    Select Case cEdMode
        Case 1
            tmpSLmode = 0
            Call CDrawer.Draw_Point(intX, intY)
            IsChanged = True
            Call SetMainCaption
        Case 2
            tmpSLmode = 1
        Case 3
            If Option1(0).Value = True Then
                tmpSLmode = 0
                Call CDrawer.ExChColor_Point(intX, intY, ExChColorVar(0), ExChColorVar(1))
                IsChanged = True
                Call SetMainCaption
            Else
                tmpSLmode = 1
            End If
        Case 4
            tmpSLmode = 1
        Case 5
            If CPBflag = 0 Then
                tmpSLmode = 1
            Else
                If Check1.Value = 1 Then
                    Call CDrawer.CPB_Past_DeBackGround(intX, intY)
                Else
                    Call CDrawer.CPB_Past(intX, intY)
                End If
                IsChanged = True
                Call SetMainCaption
                tmpSLmode = 3
            End If
    
    End Select
    If tmpSLmode = 1 Then
        SL.StartPoint_X = intX
        SL.StartPoint_Y = intY
        SL.EndPoint_X = intX
        SL.EndPoint_Y = intY
        SL.DrawSelect
        Label6.Caption = "(" & intX + 1 & "," & intY + 1 & ")-(" & intX + 1 & "," & intY + 1 & ")"
    Else
        Label6.Caption = ""
    End If

    'Picture1.Refresh
Exit Sub
out:
Debug.Print "Picture1_MouseDown Error Out : " & Err.Description
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo out
    
    intX = Fix(X)
    intY = Fix(Y)
    If FixMouseMove(0) = intX And FixMouseMove(1) = intY Then Exit Sub
    FixMouseMove(0) = intX
    FixMouseMove(1) = intY
    Dim tmpSLmode As Byte   '��ܮؽu���Ҧ�
    Select Case cEdMode
        Case 1
            tmpSLmode = 0
            
            If Button = 1 Then Call CDrawer.Draw_Point(intX, intY)
        Case 2
            If Button = 1 Then
                tmpSLmode = 1
            Else
                tmpSLmode = 2
            End If
        Case 3
            If Option1(0).Value = True Then
                tmpSLmode = 0
                If Button = 1 Then
                    Call CDrawer.ExChColor_Point(intX, intY, ExChColorVar(0), ExChColorVar(1))
                    IsChanged = True
                    Call SetMainCaption
                End If
            Else
                If Button = 1 Then
                    tmpSLmode = 1
                Else
                    tmpSLmode = 2
                End If
            End If
        Case 4
            If Button = 1 Then
                tmpSLmode = 1
            Else
                tmpSLmode = 2
            End If
        Case 5
            If CPBflag = 0 Then
                If Button = 1 Then
                    tmpSLmode = 1
                Else
                    tmpSLmode = 2
                End If
            Else
                tmpSLmode = 3
            End If
    
    End Select
    Select Case tmpSLmode
        Case 0
        
            SL.StartPoint_X = intX
            SL.StartPoint_Y = intY
            SL.EndPoint_X = intX
            SL.EndPoint_Y = intY
            Call SL.DrawSelect
        Case 1
            SL.EndPoint_X = intX
            SL.EndPoint_Y = intY
            Call SL.DrawSelect
            Label6.Caption = "(" & SL.StartPoint_X + 1 & "," & SL.StartPoint_Y + 1 & ")-(" & SL.EndPoint_X + 1 & "," & SL.EndPoint_Y + 1 & ")"
        Case 2
        Case 3
            SL.StartPoint_X = intX
            SL.StartPoint_Y = intY
            SL.EndPoint_X = intX + CDrawer.CPB_uX
            SL.EndPoint_Y = intY + CDrawer.CPB_uY
            Call SL.DrawSelect
            
    End Select
    Label1.Caption = "(" & intX + 1 & "," & intY + 1 & ")" '& Button
Exit Sub
out:
Debug.Print "Picture1_MouseMove Error Out : " & Err.Description
End Sub

Public Sub SetToolP(ByVal newcTI As Integer)
    Dim tmpcTI As Integer
    ToolP.Visible = False
    tmpcTI = cToolIndex
    cToolIndex = newcTI
    Set ToolP = Frame2(cToolIndex)
    ToolP.Left = Frame2(tmpcTI).Left
    ToolP.Top = Frame2(tmpcTI).Top
    ToolP.Visible = True
    
End Sub

Public Sub SetUIPos()
    
    'ToolP.Left = Picture1.Left + Picture1.Width
    If Picture1.Left + Picture1.Width < Frame4.Left + Frame4.Width Then
        ToolP.Left = Frame4.Left + Frame4.Width
    Else
        ToolP.Left = Picture1.Left + Picture1.Width
    End If
    If Picture1.Top + Picture1.Height < ToolP.Top + ToolP.Height Then
        Frame3.Top = ToolP.Top + ToolP.Height
    Else
        Frame3.Top = Picture1.Top + Picture1.Height
    End If
    If Me.WindowState = 0 Then
        If (ToolP.Left + ToolP.Width + 8) < (Frame3.Left + Frame3.Width) Then
            Me.Width = (Frame3.Left + Frame3.Width) * Screen.TwipsPerPixelX
        Else
            Me.Width = (ToolP.Left + ToolP.Width + 8) * Screen.TwipsPerPixelX
        End If
        Me.Height = (Frame3.Top + Frame3.Height + 45) * Screen.TwipsPerPixelY
        Me.Left = (Screen.Width - Me.Width) \ 2
        Me.Top = (Screen.Height - Me.Height) \ 2
    End If


End Sub



Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim tmpcptag As Byte '�ΨӪ�ܬO�_���}�ƻs�K�W
    '���w�ثe�s��Ҧ�
    cEdMode = Button.Index
    '�]�w�u���ݩʭ�
    Call SetToolP(cEdMode - 1)
    Select Case Button.Index
        Case 1
                SL.EndPoint_X = SL.StartPoint_X
                SL.EndPoint_Y = SL.StartPoint_Y
                Call SL.DrawSelect
        Case 2, 4, 5
        
        Case 5
            
        Case 3
            If Option1(0).Value = True Then
                SL.EndPoint_X = SL.StartPoint_X
                SL.EndPoint_Y = SL.StartPoint_Y
                Call SL.DrawSelect
            End If
    
    End Select
    
    If CPBflag = 1 And Button.Index <> 5 Then
        Call Command10_Click(2) '�����K�W�Ҧ�
    End If
End Sub

Public Sub EnableEdit(ByVal TorF As Boolean)
    If TorF Then
        ToolP.Enabled = True
        Toolbar1.Enabled = True
        Frame1.Enabled = True
        Frame3.Enabled = True
        Picture1.Enabled = True
        Me_Compile.Enabled = True
        Me_File_Save.Enabled = True
        Me_File_Save_As.Enabled = True
    Else
        ToolP.Enabled = False
        Toolbar1.Enabled = False
        Frame1.Enabled = False
        Frame3.Enabled = False
        Picture1.Enabled = False
        Me_Compile.Enabled = False
        Me_File_Save.Enabled = False
        Me_File_Save_As.Enabled = False
    End If
End Sub

Public Sub Refresh_ComboPage()
'Debug.Print "CDrawer.Z=" & CDrawer.Z
    Combo1.Clear
    For i = 1 To CDrawer.Z
        Combo1.AddItem "��" & i & "��"
    Next i
    Combo1.ListIndex = CDrawer.CurrentPage - 1
End Sub

Public Sub SetMainCaption()

    If CDrawer.cFilepath = "" Then
        Me.Caption = "Visual Ansi �����"
    Else
        Me.Caption = "Visual Ansi ����� - " & GetFileName(CDrawer.cFilepath)
        If IsChanged = True Then Me.Caption = Me.Caption & "*"
    End If
End Sub

Public Function GetFileName(filepath As String) As String
        Dim tmpstrA() As String
        tmpstrA = Split(filepath, "\")
        GetFileName = tmpstrA(UBound(tmpstrA))
End Function
Public Sub AskSave()
If IsChanged = True Then
        If MsgBox("�O�_�n�b�����e�x�s�{�b�o���ɮ�" & vbCrLf & CDrawer.cFilepath, 36, "����") = 6 Then
            Call Me_File_Save_Click
        End If
End If
End Sub
Private Sub OpenFile_Command(ByVal filepath As String)
On Error GoTo out
    If Dir(filepath, vbHidden Or vbDirectory Or vbReadOnly Or vbSystem) = "" Then
        MsgBox "�ɮפ��s�b", 16, "���~"
        Exit Sub
    End If
    Call CDrawer.OpenFile(filepath)
    Call Form16.ReSize_UI
    Call EnableEdit(True)
    Call Refresh_ComboPage
    Call Command2_Click(3)
    IsChanged = False
    Call SetMainCaption
    Call SetUIPos
    
Exit Sub
out:
    Debug.Print "OpenFile_Command Error Out"

End Sub
