VERSION 5.00
Begin VB.Form Form18 
   BorderStyle     =   5  '可調整工具視窗
   Caption         =   "編譯-方塊圖"
   ClientHeight    =   4110
   ClientLeft      =   60
   ClientTop       =   285
   ClientWidth     =   3615
   LinkTopic       =   "Form18"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4110
   ScaleWidth      =   3615
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所屬視窗中央
   Begin VB.CommandButton Command1 
      Caption         =   "執行"
      Height          =   345
      Left            =   450
      TabIndex        =   10
      Top             =   3675
      Width           =   1140
   End
   Begin VB.CommandButton Command2 
      Caption         =   "關閉"
      Height          =   330
      Left            =   1950
      TabIndex        =   9
      Top             =   3675
      Width           =   1230
   End
   Begin VB.Frame Frame1 
      Caption         =   "輸出"
      Height          =   930
      Left            =   75
      TabIndex        =   7
      Top             =   2505
      Width           =   3450
      Begin VB.CheckBox Check1 
         Caption         =   "存成文字檔(ANS)"
         Height          =   180
         Index           =   1
         Left            =   165
         TabIndex        =   15
         Top             =   615
         Width           =   1980
      End
      Begin VB.CheckBox Check1 
         Caption         =   "複製到剪貼簿"
         Height          =   180
         Index           =   0
         Left            =   165
         TabIndex        =   8
         Top             =   285
         Value           =   1  '核取
         Width           =   1620
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "選項"
      Height          =   2295
      Left            =   90
      TabIndex        =   0
      Top             =   120
      Width           =   3450
      Begin VB.CheckBox Check4 
         Caption         =   "保護行末空白色彩"
         Height          =   180
         Left            =   390
         TabIndex        =   14
         Top             =   1725
         Value           =   1  '核取
         Width           =   1770
      End
      Begin VB.CheckBox Check3 
         Caption         =   "為動畫加入頁數"
         Height          =   210
         Left            =   390
         TabIndex        =   13
         Top             =   2010
         Width           =   2505
      End
      Begin VB.CheckBox Check2 
         Caption         =   "避免游標引起的色彩喪失"
         Height          =   180
         Left            =   390
         TabIndex        =   11
         Top             =   1260
         Value           =   1  '核取
         Width           =   2580
      End
      Begin VB.TextBox Text1 
         Height          =   240
         Left            =   1260
         TabIndex        =   4
         Text            =   "---"
         Top             =   915
         Width           =   1740
      End
      Begin VB.OptionButton Option1 
         Caption         =   "全部"
         Height          =   240
         Index           =   0
         Left            =   375
         TabIndex        =   3
         Top             =   525
         Value           =   -1  'True
         Width           =   705
      End
      Begin VB.OptionButton Option1 
         Caption         =   "本頁"
         Height          =   240
         Index           =   1
         Left            =   1140
         TabIndex        =   2
         Top             =   510
         Width           =   690
      End
      Begin VB.OptionButton Option1 
         Caption         =   "選取範圍"
         Height          =   240
         Index           =   2
         Left            =   1860
         TabIndex        =   1
         Top             =   510
         Width           =   1035
      End
      Begin VB.Label Label3 
         Caption         =   "(這個選項會使彩色碼變肥些)"
         Height          =   225
         Left            =   660
         TabIndex        =   12
         Top             =   1470
         Width           =   2460
      End
      Begin VB.Label Label1 
         Caption         =   "動畫重複行: "
         Height          =   210
         Left            =   225
         TabIndex        =   6
         Top             =   945
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "範圍："
         Height          =   210
         Left            =   240
         TabIndex        =   5
         Top             =   255
         Width           =   540
      End
   End
End
Attribute VB_Name = "Form18"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Me.Hide
    Form12.Show
    Load Form6
    Dim tmpAT
    Dim tmpInt As Integer
    Form6.Visible = False
    Me.Visible = False
    Call Form14.CDrawer.SetByteArray
    If Option1(0).Value = True Then
        'tmpAT = Form14.CDrawer.GetAnsiText_All(Check2.value, Check3.value, Check4.value)
        Call Form14.CDrawer.GetAnsiText_All(Check2.Value, Check3.Value, Check4.Value)
    End If
    
    If Option1(1).Value = True Then

        'tmpAT = Form14.CDrawer.GetAnsiText_Area(0, 0, Form14.CDrawer.W - 1, Form14.CDrawer.H - 1, Form14.CDrawer.CurrentPage, Check2.value, Check4.value)
        Call Form14.CDrawer.GetAnsiText_Area(0, 0, Form14.CDrawer.W - 1, Form14.CDrawer.H - 1, Form14.CDrawer.CurrentPage, Check2.Value, Check4.Value)

    End If
    If Option1(2).Value = True Then
        tmpInt = Form14.SL.StartPoint_Y - Form14.SL.EndPoint_Y + 1
        Call Form14.SL.CorrectPoints
        If (tmpInt Mod 2) <> 0 Then
            If Form14.SL.EndPoint_Y > Form14.CDrawer.H - 1 Then
                Form14.SL.EndPoint_Y = Form14.SL.EndPoint_Y - 1
            Else
                Form14.SL.EndPoint_Y = Form14.SL.EndPoint_Y + 1
            End If
            Call Form14.SL.DrawSelect
            MsgBox "高必須要是偶數 故修正為(" & Form14.SL.StartPoint_X & "," & Form14.SL.StartPoint_Y & ")-(" & Form14.SL.EndPoint_X & "," & Form14.SL.EndPoint_Y & ")"
        End If

            'tmpAT = Form14.CDrawer.GetAnsiText_Area(Form14.SL.StartPoint_X, Form14.SL.StartPoint_Y, Form14.SL.EndPoint_X, Form14.SL.EndPoint_Y, Form14.CDrawer.CurrentPage, Check2.value, Check4.value)
            Call Form14.CDrawer.GetAnsiText_Area(Form14.SL.StartPoint_X, Form14.SL.StartPoint_Y, Form14.SL.EndPoint_X, Form14.SL.EndPoint_Y, Form14.CDrawer.CurrentPage, Check2.Value, Check4.Value)

        'tmpAT = Form14.CDrawer.GetAnsiText_Area(Form14.SL.StartPoint_X, Form14.SL.StartPoint_Y, Form14.SL.EndPoint_X, Form14.SL.EndPoint_Y, Form14.CDrawer.CurrentPage)
    End If
    If Check1(0).Value = 1 Then
        Call Form14.CDrawer.BA_ClipBoard_Copy(Me.hwnd)   '將完成的內容複製到剪貼簿
        MsgBox "已經複製到剪貼簿，如果Ctrl+V失效 可用Shift+Insert取代", 64, "完成"
    End If
    If Check1(1).Value = 1 Then
        
        Dim outfile As Integer
        Dim newfilename As String
        Form14.CDialog1.DialogTitle = "另存新檔"
        Form14.CDialog1.Filter = "*.ans(ANSI彩色檔)|*.ans"
        If Form14.CDrawer.cFilepath = "" Then
            Form14.CDialog1.FileName = "未命名.ans"
        Else
            Form14.CDialog1.FileName = Left(Form14.CDrawer.cFilepath, Len(Form14.CDrawer.cFilepath) - 4) & ".ans"
        End If
        Form14.CDialog1.ShowSave
        
        If Dir(Form14.CDialog1.FileName) <> "" Then
            If MsgBox("這個檔案已經存在,你確定要覆蓋它嗎?", vbOKCancel, "檔案已存在") = vbNo Then Exit Sub
        End If
            
        If Form14.CDialog1.FileName <> "" Then
            Call Form14.CDrawer.BA_SaveAnsiFile(Form14.CDialog1.FileName)
        End If
    End If
    Unload Form12
    'Form6.Text1.text = tmpAT
    'Form6.Visible = True
    Me.Show
    'Unload Me
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

