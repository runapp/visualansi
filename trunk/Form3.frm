VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form3 
   BorderStyle     =   4  '單線固定工具視窗
   Caption         =   "Ansi List 編輯器"
   ClientHeight    =   3180
   ClientLeft      =   1425
   ClientTop       =   480
   ClientWidth     =   5205
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3180
   ScaleWidth      =   5205
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Left            =   5580
      Top             =   885
   End
   Begin VB.ListBox List1 
      Appearance      =   0  '平面
      BackColor       =   &H00000000&
      ForeColor       =   &H00C0C0C0&
      Height          =   2550
      Left            =   6210
      TabIndex        =   3
      Top             =   1230
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Frame Frame1 
      Height          =   3075
      Left            =   3435
      TabIndex        =   0
      Top             =   -15
      Width           =   1635
      Begin VB.CommandButton Command5 
         Caption         =   "關閉"
         Height          =   285
         Left            =   180
         TabIndex        =   13
         Top             =   2670
         Width           =   1230
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  '置中對齊
         Appearance      =   0  '平面
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "細明體"
            Size            =   14.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   420
         Left            =   555
         TabIndex        =   11
         Top             =   495
         Width           =   480
      End
      Begin VB.CommandButton Command1 
         Caption         =   "→"
         Height          =   330
         Index           =   3
         Left            =   1065
         Style           =   1  '圖片外觀
         TabIndex        =   10
         Top             =   525
         Width           =   360
      End
      Begin VB.CommandButton Command1 
         Caption         =   "←"
         Height          =   330
         Index           =   2
         Left            =   195
         Style           =   1  '圖片外觀
         TabIndex        =   9
         Top             =   555
         Width           =   330
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Left            =   1095
         MaxLength       =   1
         TabIndex        =   7
         Top             =   1380
         Width           =   345
      End
      Begin VB.CommandButton Command4 
         Caption         =   "取代為"
         Enabled         =   0   'False
         Height          =   285
         Left            =   165
         TabIndex        =   6
         Top             =   1440
         Width           =   795
      End
      Begin VB.CommandButton Command3 
         Caption         =   "刪除"
         Height          =   285
         Left            =   165
         TabIndex        =   5
         Top             =   1860
         Width           =   1260
      End
      Begin VB.CommandButton Command2 
         Caption         =   "儲存套用"
         Height          =   315
         Left            =   180
         TabIndex        =   4
         Top             =   2235
         Width           =   1245
      End
      Begin VB.CommandButton Command1 
         Caption         =   "↓"
         Height          =   330
         Index           =   1
         Left            =   615
         Style           =   1  '圖片外觀
         TabIndex        =   2
         Top             =   930
         Width           =   360
      End
      Begin VB.CommandButton Command1 
         Caption         =   "↑"
         Height          =   330
         Index           =   0
         Left            =   615
         Style           =   1  '圖片外觀
         TabIndex        =   1
         Top             =   150
         Width           =   360
      End
      Begin VB.Label Label1 
         Caption         =   "移動"
         Height          =   225
         Left            =   135
         TabIndex        =   12
         Top             =   240
         Width           =   420
      End
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   3000
      Left            =   60
      TabIndex        =   8
      Top             =   75
      Width           =   3330
      _ExtentX        =   5874
      _ExtentY        =   5292
      _Version        =   393216
      Rows            =   5
      Cols            =   10
      FixedRows       =   0
      FixedCols       =   0
      AllowBigSelection=   0   'False
      FocusRect       =   2
      GridLinesFixed  =   1
      ScrollBars      =   2
      AllowUserResizing=   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "細明體"
         Size            =   11.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MoveMode As Byte
Dim KeepMoving As Boolean
Private Sub Command1_Click(Index As Integer)
Dim row As Integer
Dim col As Integer
Dim KeyStr As String
MSFlexGrid1.SetFocus
row = MSFlexGrid1.RowSel
col = MSFlexGrid1.ColSel

Select Case Index
    Case 0

        ExChGrid MSFlexGrid1, row, col, row - 1, col
        KeyStr = "{UP}"
    Case 1
        ExChGrid MSFlexGrid1, row, col, row + 1, col
        KeyStr = "{Down}"
    Case 2
        ExChGrid MSFlexGrid1, row, col, row, col - 1
        
        KeyStr = "{Left}"
    Case 3
        ExChGrid MSFlexGrid1, row, col, row, col + 1
        KeyStr = "{Right}"
End Select
SendKeys KeyStr

End Sub

Private Sub Command1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'Select Case Index
'    Case 0
'        MoveMode = 1
'    Case 1
'        MoveMode = 2
'End Select
'Call Timer1_Timer
'Timer1.Interval = 200

End Sub

Private Sub Command1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'MoveMode = 0
'Timer1.Interval = 0
End Sub

Private Sub Command2_Click()
'Call SaveList
MSFlexGrid1.SetFocus
Call SaveGrid
Call LoadAnsi(App.Path & "\Ansi.txt", Form1.MSFlexGrid1)
End Sub

Private Sub Command3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'MoveMode = 3
'Call Timer1_Timer
'Timer1.Interval = 190
MSFlexGrid1.SetFocus
MSFlexGrid1.text = " "
End Sub

Private Sub Command3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'MoveMode = 0
'Timer1.Interval = 0
End Sub

Private Sub Command4_Click()
MSFlexGrid1.SetFocus
'List1.AddItem Text1.text
'List1.ListIndex = List1.ListCount - 1
MSFlexGrid1.text = Text1.text
End Sub





Private Sub Command5_Click()
    Unload Me
End Sub

Private Sub Form_Load()
'Call LoadAnsiLs(App.Path & "\ansi.txt", List1)
For i = 0 To 9
     MSFlexGrid1.ColWidth(i) = 300


Next i
Call LoadAnsi(App.Path & "\Ansi.txt", Form3.MSFlexGrid1)
End Sub

Public Sub SaveList()
Dim tempfile As Integer
Dim tempstr As String
tempfile = 12


Open App.Path & "\ansi.txt" For Binary As #tempfile

    For i = 0 To (List1.ListCount - 1)
        tempstr = tempstr & List1.List(i)

    Next i
Put #tempfile, 1, tempstr


Close #tempfile





End Sub


Private Sub MSFlexGrid1_Click()
Text2.text = MSFlexGrid1.text
End Sub

Private Sub Text1_Change()
If Trim(Text1.text) = "" Then
    
    Command4.Enabled = False
Else
    Command4.Enabled = True

End If
End Sub


Private Sub Timer1_Timer()
Dim x1 As Integer

x1 = List1.ListIndex



Select Case MoveMode
    Case 1
        Call ExChList(List1, x1, x1 - 1)
        List1.ListIndex = x1 - 1
    Case 2
                Call ExChList(List1, x1, x1 + 1)
        List1.ListIndex = x1 + 1
    Case 3
    On Error Resume Next
        List1.RemoveItem x1
        List1.ListIndex = x1
End Select


End Sub
Public Sub SaveGrid()
Dim tempfile As Integer
Dim tempstr As String
tempfile = 12
Dim counter As Integer
Kill App.Path & "\ansi.txt"
Open App.Path & "\ansi.txt" For Binary As #tempfile

counter = 0

Do
    If MSFlexGrid1.TextMatrix(counter \ 10, counter Mod 10) = "" Then Exit Do
    tempstr = tempstr & MSFlexGrid1.TextMatrix(counter \ 10, counter Mod 10)
    counter = counter + 1
    DoEvents
Loop

Put #tempfile, 1, tempstr


Close #tempfile





End Sub
