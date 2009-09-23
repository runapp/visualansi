VERSION 5.00
Begin VB.Form Form22 
   BorderStyle     =   5  '可調整工具視窗
   Caption         =   "輸出成HTML網頁"
   ClientHeight    =   3180
   ClientLeft      =   60
   ClientTop       =   285
   ClientWidth     =   5985
   LinkTopic       =   "Form22"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3180
   ScaleWidth      =   5985
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所屬視窗中央
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   405
      Left            =   5430
      TabIndex        =   24
      Top             =   3270
      Width           =   405
   End
   Begin VB.TextBox html_text 
      Height          =   2340
      Left            =   105
      MultiLine       =   -1  'True
      ScrollBars      =   3  '兩者皆有
      TabIndex        =   22
      Text            =   "Form22.frx":0000
      Top             =   3405
      Width           =   5550
   End
   Begin VB.Frame Frame4 
      Caption         =   "網頁設定"
      Height          =   1560
      Left            =   3330
      TabIndex        =   4
      Top             =   210
      Width           =   2295
      Begin VB.TextBox text_font_size_unit 
         Alignment       =   2  '置中對齊
         Height          =   270
         Left            =   1410
         TabIndex        =   23
         Text            =   "pt"
         Top             =   570
         Width           =   360
      End
      Begin VB.TextBox text_font_size 
         Alignment       =   1  '靠右對齊
         Height          =   270
         Left            =   870
         TabIndex        =   16
         Text            =   "18"
         Top             =   570
         Width           =   435
      End
      Begin VB.TextBox text_title 
         Height          =   270
         Left            =   540
         TabIndex        =   6
         Text            =   "BBS Movie"
         Top             =   225
         Width           =   1650
      End
      Begin VB.Label Label4 
         Caption         =   "字體大小"
         Height          =   225
         Left            =   120
         TabIndex        =   15
         Top             =   615
         Width           =   750
      End
      Begin VB.Label Label1 
         Caption         =   "標題"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   285
         Width           =   495
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Javascript功能"
      Height          =   1230
      Left            =   75
      TabIndex        =   3
      Top             =   1215
      Width           =   3135
      Begin VB.TextBox text_forcetime 
         Alignment       =   1  '靠右對齊
         Height          =   270
         Left            =   1980
         TabIndex        =   20
         Text            =   "0.6"
         Top             =   750
         Width           =   405
      End
      Begin VB.OptionButton withtimeline 
         Caption         =   "強制每頁時間為"
         Height          =   180
         Index           =   1
         Left            =   420
         TabIndex        =   19
         Top             =   795
         Width           =   1590
      End
      Begin VB.OptionButton withtimeline 
         Caption         =   "使用時間軸時間"
         Height          =   180
         Index           =   0
         Left            =   420
         TabIndex        =   18
         Top             =   555
         Value           =   -1  'True
         Width           =   1560
      End
      Begin VB.CheckBox autoplay_ctl 
         Caption         =   "自動播放功能"
         Height          =   195
         Left            =   195
         TabIndex        =   17
         Top             =   285
         Value           =   1  '核取
         Width           =   1515
      End
      Begin VB.Label Label6 
         Caption         =   "秒"
         Height          =   240
         Left            =   2445
         TabIndex        =   21
         Top             =   795
         Width           =   285
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "輸出範圍"
      Height          =   945
      Left            =   75
      TabIndex        =   2
      Top             =   210
      Width           =   3135
      Begin VB.OptionButton range_opt 
         Caption         =   "全部"
         Height          =   240
         Index           =   0
         Left            =   255
         TabIndex        =   12
         Top             =   285
         Value           =   -1  'True
         Width           =   705
      End
      Begin VB.OptionButton range_opt 
         Caption         =   "本頁"
         Height          =   240
         Index           =   1
         Left            =   960
         TabIndex        =   11
         Top             =   270
         Visible         =   0   'False
         Width           =   705
      End
      Begin VB.OptionButton range_opt 
         Caption         =   "選取範圍"
         Height          =   240
         Index           =   2
         Left            =   1710
         TabIndex        =   10
         Top             =   270
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox Text2 
         Height          =   270
         Left            =   1185
         TabIndex        =   9
         Text            =   "1"
         Top             =   540
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.TextBox Text3 
         Height          =   270
         Left            =   1740
         TabIndex        =   8
         Text            =   "2"
         Top             =   540
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.OptionButton range_opt 
         Caption         =   "特定頁"
         Height          =   240
         Index           =   3
         Left            =   255
         TabIndex        =   7
         Top             =   570
         Visible         =   0   'False
         Width           =   870
      End
      Begin VB.Label Label2 
         Caption         =   "~"
         Height          =   225
         Left            =   1590
         TabIndex        =   14
         Top             =   585
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.Label Label3 
         Caption         =   "頁"
         Height          =   195
         Left            =   2250
         TabIndex        =   13
         Top             =   585
         Visible         =   0   'False
         Width           =   300
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "關閉"
      Height          =   375
      Index           =   1
      Left            =   3240
      TabIndex        =   1
      Top             =   2640
      Width           =   1230
   End
   Begin VB.CommandButton Command1 
      Caption         =   "輸出"
      Height          =   375
      Index           =   0
      Left            =   1320
      TabIndex        =   0
      Top             =   2640
      Width           =   1230
   End
   Begin VB.Label Label5 
      Caption         =   "注意：不是所有的字體大小都能讓字跟字之間緊密結合，預設為18pt。"
      Height          =   630
      Left            =   3330
      TabIndex        =   25
      Top             =   1800
      Width           =   2385
   End
End
Attribute VB_Name = "Form22"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub autoplay_ctl_Click()
    If autoplay_ctl.Value Then
        withtimeline(0).Enabled = True
        withtimeline(1).Enabled = True
    Else
        withtimeline(0).Enabled = False
        withtimeline(1).Enabled = False
    
    End If
End Sub

Private Sub Command1_Click(Index As Integer)
On Error GoTo out
    If Index = 0 Then
        Me.Hide
        Form12.Show
        Dim tmp_html As String
        Dim forcetime As Single
        If withtimeline(0).Value Then
            forcetime = 0
        Else
            forcetime = Val(text_forcetime.text)
        End If
        Call Form1.CreatHTML_VAA(tmp_html, 1, UBound(Arrf, 3), text_title.text, autoplay_ctl.Value, forcetime)
        'html_text.text = tmp_html
        Unload Form12
            Dim outfile As Integer
            Dim newfilename As String
            Form1.CDialog1.DialogTitle = "另存新檔"
            Form1.CDialog1.Filter = "*.html(網頁)|*.html"
            If OFP.FilePath <> "" Then
                Form1.CDialog1.FileName = Left(OFP.FilePath, Len(OFP.FilePath) - 4) & ".html"
            Else
                Form1.CDialog1.FileName = "ansi.html"
            End If
                
            Form1.CDialog1.ShowSave
            
            If Form1.FileSys.FileExists(Form1.CDialog1.FileName) = True Then
                If MsgBox("這個檔案已經存在,你確定要覆蓋它嗎?", vbOKCancel, "檔案已存在") = vbNo Then Exit Sub
                Kill Form1.CDialog1.FileName
            End If
            
            If Form1.CDialog1.FileName <> "" Then
                outfile = 40
                Open Form1.CDialog1.FileName For Binary As #outfile
                Put #outfile, 1, tmp_html
                Close outfile
            End If
        
    Else
        Unload Me
    
    End If
Exit Sub
out:
    'MsgBox (Err.Description)
    Unload Form12
    Unload Me
End Sub

Private Sub Command2_Click()
    Debug.Print LoadResString(101)
End Sub

Private Sub Form_Load()
    If OFP.FilePath <> "" Then
                text_title.text = Left(OFP.FilePath, Len(OFP.FilePath) - 4)
    Else
                text_title.text = "Visual Ansi"
    End If
End Sub

Private Sub text_title_GotFocus()
    'text_title.SelStart = 0
    'text_title.SelLength = Len(text_title.text)
    Text_Sel_All text_title
End Sub

Private Sub Text2_GotFocus()
    Text_Sel_All Text2
End Sub

Private Sub Text3_GotFocus()
    Text_Sel_All Text3
End Sub
