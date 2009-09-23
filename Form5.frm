VERSION 5.00
Begin VB.Form Form5 
   BorderStyle     =   4  '單線固定工具視窗
   Caption         =   "輸出成BBS彩色碼"
   ClientHeight    =   3600
   ClientLeft      =   45
   ClientTop       =   270
   ClientWidth     =   3600
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3600
   ScaleWidth      =   3600
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所屬視窗中央
   Begin VB.CommandButton Command3 
      Caption         =   "執行"
      Height          =   315
      Left            =   450
      TabIndex        =   12
      Top             =   3045
      Width           =   1230
   End
   Begin VB.CommandButton Command2 
      Caption         =   "關閉"
      Height          =   315
      Left            =   1920
      TabIndex        =   5
      Top             =   3045
      Width           =   1230
   End
   Begin VB.CommandButton Command1 
      Caption         =   "開始編譯"
      Height          =   345
      Left            =   1110
      TabIndex        =   4
      Top             =   3420
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Frame Frame2 
      Caption         =   "選項"
      Height          =   1755
      Left            =   90
      TabIndex        =   2
      Top             =   75
      Width           =   3450
      Begin VB.OptionButton Option1 
         Caption         =   "特定頁"
         Height          =   240
         Index           =   3
         Left            =   750
         TabIndex        =   19
         Top             =   630
         Width           =   870
      End
      Begin VB.TextBox Text3 
         Height          =   270
         Left            =   2235
         TabIndex        =   17
         Text            =   "2"
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox Text2 
         Height          =   270
         Left            =   1680
         TabIndex        =   15
         Text            =   "1"
         Top             =   600
         Width           =   360
      End
      Begin VB.CheckBox Check3 
         Caption         =   "動畫重複行"
         Height          =   180
         Index           =   1
         Left            =   225
         TabIndex        =   14
         Top             =   1005
         Value           =   1  '核取
         Width           =   1200
      End
      Begin VB.CheckBox Check3 
         Caption         =   "加入時間軸(僅限pmore動畫支援)"
         Height          =   180
         Index           =   0
         Left            =   225
         TabIndex        =   13
         Top             =   1305
         Width           =   2925
      End
      Begin VB.OptionButton Option1 
         Caption         =   "選取範圍"
         Height          =   240
         Index           =   2
         Left            =   2205
         TabIndex        =   10
         Top             =   330
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "本頁"
         Height          =   240
         Index           =   1
         Left            =   1455
         TabIndex        =   9
         Top             =   330
         Width           =   840
      End
      Begin VB.OptionButton Option1 
         Caption         =   "全部"
         Height          =   240
         Index           =   0
         Left            =   750
         TabIndex        =   8
         Top             =   345
         Value           =   -1  'True
         Width           =   705
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Left            =   1440
         TabIndex        =   7
         Text            =   "---"
         Top             =   960
         Width           =   1710
      End
      Begin VB.CheckBox Check2 
         Caption         =   "不含彩色碼"
         Height          =   180
         Left            =   2445
         TabIndex        =   6
         Top             =   60
         Visible         =   0   'False
         Width           =   1395
      End
      Begin VB.Label Label3 
         Caption         =   "頁"
         Height          =   195
         Left            =   2745
         TabIndex        =   18
         Top             =   645
         Width           =   300
      End
      Begin VB.Label Label1 
         Caption         =   "~"
         Height          =   225
         Left            =   2085
         TabIndex        =   16
         Top             =   645
         Width           =   150
      End
      Begin VB.Label Label2 
         Caption         =   "範圍："
         Height          =   210
         Left            =   240
         TabIndex        =   11
         Top             =   360
         Width           =   540
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "輸出"
      Height          =   945
      Left            =   75
      TabIndex        =   0
      Top             =   1950
      Width           =   3465
      Begin VB.CheckBox Check1 
         Caption         =   "存成文字檔(ANS)"
         Height          =   180
         Index           =   1
         Left            =   165
         TabIndex        =   3
         Top             =   585
         Width           =   2130
      End
      Begin VB.CheckBox Check1 
         Caption         =   "複製到剪貼簿"
         Height          =   180
         Index           =   0
         Left            =   165
         TabIndex        =   1
         Top             =   315
         Value           =   1  '核取
         Width           =   1620
      End
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub Check3_Click(Index As Integer)
    If Check3(0).Value = 1 And Index = 0 Then
        Check3(1).Value = 0
        
    ElseIf Check3(1).Value = 1 And Index = 1 Then
        Check3(0).Value = 0
        
    End If
    
End Sub

Private Sub Command1_Click()

    Dim ansistr As String
    Form5.Hide
    Form12.Show
    If Check2.Value = 1 Then
        Call Form1.CreatAnsiTxt_NoColor(ansistr)
    Else
    'Debug.Print "start:" & Time

    If Option1(0).Value = True Then
        '編譯函式已經三合一
        Call Form1.CreatAnsiTxt_VAA_v3(ansistr, Text1.text)

    Else
        If Option1(1).Value = True Then
            Call Form1.CreatAnsiTxt_Area(ansistr, 0, 0, UBound(Arrf, 1), UBound(Arrf, 2))
        Else
            Dim tmpX1 As Integer, tmpX2 As Integer, tmpY1 As Integer, tmpY2 As Integer, tmpInt As Integer
            tmpX1 = Form1.SL.StartPoint_X
            tmpY1 = Form1.SL.StartPoint_Y
            tmpX2 = Form1.SL.EndPoint_X
            tmpY2 = Form1.SL.EndPoint_Y
            If tmpX1 > tmpX2 Then
                tmpInt = tmpX1
                tmpX1 = tmpX2
                tmpX2 = tmpInt
            End If
            If tmpY1 > tmpY2 Then
                tmpInt = tmpY1
                tmpY1 = tmpY2
                tmpY2 = tmpInt
            End If
            Call Form1.CreatAnsiTxt_Area(ansistr, tmpX1, tmpY1, tmpX2, tmpY2)
        End If
        
        
    End If
End If
    
    Form5.Visible = False
    If Check1(0).Value = 1 Then
        Load Form6
        Form6.Text1.text = ansistr
        Form6.Show
    End If
'Debug.Print "end:" & Time
If Check1(1).Value = 1 Then
    Dim outfile As Integer
    Dim newfilename As String
    newfilename = Left(OFP.FilePath, Len(OFP.FilePath) - 4) & ".txt"
    outfile = 40
    Open newfilename For Binary As #outfile
    
    Put #outfile, 1, ansistr
    
    
    Close outfile
End If

Unload Form12
Unload Form5
End Sub

Private Sub Command2_Click()
Unload Form5
End Sub

Private Sub Command3_Click()
On Error GoTo err_out
    If Check1(0).Value = 0 And Check1(1).Value = 0 Then Exit Sub
    Dim ansistr As String
    Form5.Hide
    Form12.Show
    If Check2.Value = 1 Then
        Call Form1.CreatAnsiTxt_NoColor(ansistr)
    Else

        If Option1(0).Value = True Then
            '編譯全部
            Call Form1.CreatAnsiTxt_VAA_v4(Text1.text, Check3(0).Value)
        ElseIf Option1(3).Value = True Then
            '編譯特定頁
            '檢查輸入的頁數值
            Dim fromPage As Integer, toPage As Integer
            fromPage = Val(Text2.text)
            toPage = Val(Text3.text)
            If fromPage < 1 Then
                MsgBox "起始頁數不得小於1", 16, "錯誤"
                GoTo out
            End If
            If fromPage < 1 Then
                MsgBox "最大頁數不得大於" & UBound(Arrf, 3), 16, "錯誤"
                GoTo out
            End If
            If fromPage > toPage Then
                MsgBox "起始頁數不得大於結束頁數", 16, "錯誤"
                GoTo out
            End If
            Call Form1.CreatAnsiTxt_VAA_v5(fromPage, toPage, Text1.text, Check3(0).Value)
        ElseIf Option1(1).Value = True Then
            
            '編譯本頁
            Call Form1.CreatAnsiTxt_Area(ansistr, 0, 0, UBound(Arrf, 1), UBound(Arrf, 2))
        Else
            Dim tmpX1 As Integer, tmpX2 As Integer, tmpY1 As Integer, tmpY2 As Integer, tmpInt As Integer
            tmpX1 = Form1.SL.StartPoint_X
            tmpY1 = Form1.SL.StartPoint_Y
            tmpX2 = Form1.SL.EndPoint_X
            tmpY2 = Form1.SL.EndPoint_Y
            If tmpX1 > tmpX2 Then
                tmpInt = tmpX1
                tmpX1 = tmpX2
                tmpX2 = tmpInt
            End If
            If tmpY1 > tmpY2 Then
                tmpInt = tmpY1
                tmpY1 = tmpY2
                tmpY2 = tmpInt
            End If
            Call Form1.CreatAnsiTxt_Area(ansistr, tmpX1, tmpY1, tmpX2, tmpY2)
        End If
        
        'Form5.Visible = False
        If Check1(0).Value = 1 Then
            Call SetClipboard_ByteArray(Me.hwnd, ByteArray)
            MsgBox "已經複製到剪貼簿" & vbCrLf & "如果Ctrl+V失效 可用Shift+Insert 取代", 64, "完成"
        End If
        'Debug.Print "end:" & Time
        If Check1(1).Value = 1 Then
        
            Dim outfile As Integer
            Dim newfilename As String
            Form1.CDialog1.DialogTitle = "另存新檔"
            Form1.CDialog1.Filter = "*.ans(ANSI彩色檔)|*.ans"
            If OFP.FilePath <> "" Then
                Form1.CDialog1.FileName = Left(OFP.FilePath, Len(OFP.FilePath) - 4) & ".ans"
            Else
                Form1.CDialog1.FileName = "未命名.ans"
            End If
                
            Form1.CDialog1.ShowSave
            
            If Form1.FileSys.FileExists(Form1.CDialog1.FileName) = True Then
                If MsgBox("這個檔案已經存在,你確定要覆蓋它嗎?", vbOKCancel, "檔案已存在") = vbNo Then Exit Sub
            End If
            
            If Form1.CDialog1.FileName <> "" Then
                outfile = 40
                Open Form1.CDialog1.FileName For Binary As #outfile
                Put #outfile, 1, ByteArray
                Close outfile
            End If
        End If
    End If
out:
    
    Form5.Show
Exit Sub
err_out:
    Unload Form12
    Form5.Show
    Debug.Print Err.Description
End Sub

Private Sub Form_Load()
    Text3.text = UBound(Arrf, 3)
    If OFP.filetype <> 2 Then
        Text1.Enabled = False
        Check3(1).Enabled = False
        Check3(0).Enabled = False

    Else
        Text1.Enabled = True
        Check3(1).Enabled = True
        Check3(0).Enabled = True
    End If
    
    
End Sub

