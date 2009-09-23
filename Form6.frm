VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Begin VB.Form Form6 
   Caption         =   "編譯-輸出"
   ClientHeight    =   4635
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11325
   Icon            =   "Form6.frx":0000
   LinkTopic       =   "Form6"
   ScaleHeight     =   4635
   ScaleWidth      =   11325
   StartUpPosition =   3  '系統預設值
   Begin VB.Frame Frame1 
      Caption         =   "尋找可能的錯誤(此功能仍測試中)"
      Height          =   600
      Left            =   3090
      TabIndex        =   4
      Top             =   15
      Width           =   5070
      Begin VB.ComboBox Combo1 
         Appearance      =   0  '平面
         Height          =   300
         Left            =   2865
         Style           =   2  '單純下拉式
         TabIndex        =   6
         Top             =   240
         Width           =   1785
      End
      Begin VB.CommandButton Command3 
         Caption         =   "檢查錯誤"
         Height          =   315
         Left            =   210
         TabIndex        =   5
         Top             =   210
         Width           =   1440
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   255
         Left            =   1830
         TabIndex        =   7
         Top             =   270
         Width           =   1290
      End
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   255
      Left            =   5505
      TabIndex        =   3
      Top             =   105
      Width           =   1050
   End
   Begin MSComDlg.CommonDialog CDialog1 
      Left            =   9345
      Top             =   30
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.CommandButton Command2 
      Caption         =   "存成檔案"
      Height          =   315
      Left            =   1620
      TabIndex        =   2
      Top             =   165
      Width           =   1275
   End
   Begin VB.CommandButton Command1 
      Caption         =   "複製下來"
      Height          =   315
      Left            =   150
      TabIndex        =   0
      Top             =   165
      Width           =   1305
   End
   Begin RichTextLib.RichTextBox Text1 
      Height          =   3930
      Left            =   15
      TabIndex        =   1
      Top             =   645
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   6932
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      DisableNoScroll =   -1  'True
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"Form6.frx":1CFA
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim errorlog() As Variant   '紀錄錯誤 (0)紀錄錯誤個數



Private Sub Combo1_Click()
    'Text1.Find errorlog(Combo1.ListIndex + 1)(2), errorlog(Combo1.ListIndex + 1)(1)
    Text1.SetFocus
    Text1.SelStart = errorlog(Combo1.ListIndex + 1)(1)
    Text1.SelLength = Len(errorlog(Combo1.ListIndex + 1)(2))
End Sub

Private Sub Command1_Click()
Clipboard.SetText (Text1.text)
End Sub

Private Sub Command2_Click()
    On Error GoTo out
    CDialog1.DialogTitle = "存成檔案"
    CDialog1.Filter = "ANSI彩色檔案(*.ans)|*.ans|純文字檔(*.txt)|*.txt|任何檔案(*.*)|*.*"
    CDialog1.FileName = ""
    CDialog1.ShowSave
    If Dir(CDialog1.FileName, vbHidden Or vbDirectory Or vbReadOnly Or vbSystem) <> "" Then
        If MsgBox("檔案已經存在 請問你要覆蓋他嗎?", 49, "檔案已存在") = vbCancel Then Exit Sub
        Kill CDialog1.FileName
    End If
    Dim outfile As Integer
    outfile = 40
    Open CDialog1.FileName For Binary As #outfile
        Put #outfile, 1, Text1.text
    Close outfile
    Exit Sub
out:
End Sub

Private Sub Command3_Click()
    Combo1.Clear
    Dim tmpVArr() As String, tmpInt As Integer, tmpLong As Long
    Dim erroritem(2) As Variant '0 為行數 1為累計字數 2為內容
    
    ReDim errorlog(0)
    
    tmpVArr = Split(Text1.text, vbCrLf)
    tmpInt = UBound(tmpVArr)
    tmpLong = 0
    For i = 0 To tmpInt
        
        If LenB(StrConv(tmpVArr(i), vbFromUnicode)) > 274 Then
            errorlog(0) = errorlog(0) + 1
            ReDim Preserve errorlog(errorlog(0))
            erroritem(0) = i + 1
            erroritem(1) = tmpLong
            erroritem(2) = tmpVArr(i)
            errorlog(errorlog(0)) = erroritem
            Debug.Print "Line " & i + 1 & " is out of range"
        End If
        tmpLong = tmpLong + Len(tmpVArr(i)) + 2 '+2為CR LF
    Next i
    Label1.Caption = "找到" & Val(errorlog(0)) & "個錯誤"
    If errorlog(0) > 0 Then
        
        'Text1.Find errorlog(1)
        For i = 1 To errorlog(0)
            Combo1.AddItem "錯誤" & i & ".Line " & errorlog(i)(0) & "=" & LenB(StrConv(errorlog(i)(2), vbFromUnicode))
        Next i
        Combo1.ListIndex = 0
    End If
    Call CheckErrExit
    
End Sub

Private Sub Command4_Click()
Debug.Print Text1.Find(vbCrLf)
End Sub

Private Sub Form_Load()
    ReDim errorlog(0)
    Text1.RightMargin = Screen.Width * 2
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Text1.Width = Form6.Width - 150
    Text1.Height = Form6.Height - 1100
End Sub


Public Sub CheckErrExit()
    If errorlog(0) = 0 Then
        Combo1.Enabled = False
        
    Else
        Combo1.Enabled = True
    End If
End Sub

Private Sub Text1_Change()
    errorlog(0) = 0
    Call CheckErrExit
    Label1.Caption = ""
End Sub

