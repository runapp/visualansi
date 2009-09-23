VERSION 5.00
Begin VB.Form Form15 
   BorderStyle     =   4  '單線固定工具視窗
   Caption         =   "建立新檔"
   ClientHeight    =   3705
   ClientLeft      =   45
   ClientTop       =   270
   ClientWidth     =   4335
   LinkTopic       =   "Form15"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3705
   ScaleWidth      =   4335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所屬視窗中央
   Begin VB.CommandButton Command1 
      Caption         =   "取消"
      Height          =   285
      Index           =   1
      Left            =   2430
      TabIndex        =   16
      Top             =   3300
      Width           =   1005
   End
   Begin VB.CommandButton Command1 
      Caption         =   "確 定"
      Height          =   285
      Index           =   0
      Left            =   675
      TabIndex        =   15
      Top             =   3315
      Width           =   1005
   End
   Begin VB.Frame Frame2 
      Caption         =   "屬性"
      Height          =   2610
      Left            =   30
      TabIndex        =   2
      Top             =   600
      Width           =   4275
      Begin VB.TextBox Text3 
         Alignment       =   1  '靠右對齊
         Height          =   240
         Left            =   3045
         TabIndex        =   18
         Text            =   "79"
         Top             =   510
         Width           =   405
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  '靠右對齊
         Height          =   270
         Left            =   2460
         TabIndex        =   13
         Text            =   "1"
         Top             =   1500
         Width           =   540
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  '靠右對齊
         Height          =   255
         Index           =   1
         Left            =   3660
         TabIndex        =   10
         Text            =   "46"
         Top             =   1065
         Width           =   465
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  '靠右對齊
         Height          =   255
         Index           =   0
         Left            =   2880
         TabIndex        =   9
         Text            =   "80"
         Top             =   1065
         Width           =   450
      End
      Begin VB.OptionButton Option3 
         Caption         =   "自訂"
         Height          =   240
         Index           =   3
         Left            =   2130
         TabIndex        =   8
         Top             =   1080
         Width           =   705
      End
      Begin VB.OptionButton Option3 
         Caption         =   "動畫頁(               X 42)"
         Height          =   240
         Index           =   2
         Left            =   2115
         TabIndex        =   6
         Top             =   525
         Width           =   2115
      End
      Begin VB.OptionButton Option3 
         Caption         =   "簽名檔(79 X 12)"
         Height          =   240
         Index           =   1
         Left            =   465
         TabIndex        =   5
         Top             =   1110
         Width           =   1560
      End
      Begin VB.OptionButton Option3 
         Caption         =   "全頁(79 X 46)"
         Height          =   240
         Index           =   0
         Left            =   465
         TabIndex        =   4
         Top             =   555
         Value           =   -1  'True
         Width           =   1425
      End
      Begin VB.Label Label4 
         Caption         =   "頁"
         Height          =   180
         Left            =   3120
         TabIndex        =   14
         Top             =   1530
         Width           =   210
      End
      Begin VB.Label Label1 
         Caption         =   "預設頁數(之後可再調整)"
         Height          =   180
         Index           =   1
         Left            =   450
         TabIndex        =   12
         Top             =   1575
         Width           =   1995
      End
      Begin VB.Label Label3 
         Caption         =   "X"
         Height          =   225
         Left            =   3420
         TabIndex        =   11
         Top             =   1110
         Width           =   135
      End
      Begin VB.Label Label2 
         Caption         =   "p.s 兩格高度=實際一行高 故高度必定為偶數"
         Height          =   255
         Left            =   495
         TabIndex        =   7
         Top             =   1995
         Width           =   3495
      End
      Begin VB.Label Label1 
         Caption         =   "畫布大小(寬 X 高)"
         Height          =   180
         Index           =   0
         Left            =   195
         TabIndex        =   3
         Top             =   285
         Width           =   1440
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "檔案類型"
      Height          =   525
      Left            =   30
      TabIndex        =   0
      Top             =   45
      Width           =   4260
      Begin VB.OptionButton Option1 
         Caption         =   "動畫"
         Height          =   225
         Index           =   1
         Left            =   2325
         TabIndex        =   17
         Top             =   210
         Width           =   675
      End
      Begin VB.OptionButton Option1 
         Caption         =   "一般"
         Height          =   225
         Index           =   0
         Left            =   855
         TabIndex        =   1
         Top             =   225
         Value           =   -1  'True
         Width           =   675
      End
   End
End
Attribute VB_Name = "Form15"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click(Index As Integer)
    If Index = 1 Then
        Unload Me
    Else
        '宣告變數
        Dim tmpW As Integer, tmpH As Integer, tmpZ As Integer
        tmpZ = Fix(Text2.text)  '取得頁數
        
        If tmpZ <= 0 Then
            If MsgBox("頁數必須為大於0的整數 已自動調整為1", 49, "數值不合") = vbCancel Then Exit Sub
            tmpZ = 1
        End If
        If Option1(0).value = True Then
            If Option3(0).value = True Then '全頁
                tmpW = 79
                tmpH = 46
            End If
            
            If Option3(1).value = True Then
                tmpW = 79
                tmpH = 12
            End If
           
            If Option3(2).value = True Then
                tmpW = Val(Text3.text)
                tmpH = 42
            End If
            
            If Option3(3).value = True Then
                tmpW = Fix(Text1(0).text)
                tmpH = Fix(Text1(1).text)
                If tmpW <= 0 Or tmpH Mod 2 = 1 Or tmpH <= 0 Then
                    If tmpW < 0 Then tmpW = Abs(tmpW)
                    If tmpH < 0 Then tmpH = Abs(tmpH)
                    If tmpW = 0 Then tmpW = 1
                    If tmpH = 0 Then tmpH = 2
                    If tmpH Mod 2 = 1 Then tmpH = tmpH + 1
                    If MsgBox("寬高必須大於零 且高必須要是偶數" & vbCrLf & "已修正為 " & tmpW & " X " & tmpH, 49, "數值不合") = vbCancel Then Exit Sub
                
                End If
            End If
            
            If tmpW >= 80 Then
                If MsgBox("寬的建議值在79以下,超過容易造成斷行,你確定要繼續嗎?", 48 + 1, "提醒") = vbCancel Then GoTo out
            End If
            If tmpW < 1 Then
                MsgBox "寬不得為0!!!", 48, "不能這樣"
                GoTo out
            End If
            
            Call Form14.CDrawer.NewFile(tmpW, tmpH, tmpZ, 4)
            Call Form14.SetUIPos
            
            Call Form16.ReSize_UI
            Call Form14.EnableEdit(True)
            Call Form14.Refresh_ComboPage
            Call Form14.Command2_Click(3)
            Unload Me
        Else
            tmpW = Val(Text3.text)
            tmpH = 42
            If tmpW >= 80 Then
                If MsgBox("寬的建議值在79以下,超過容易造成斷行,你確定要繼續嗎?", 48 + 1, "提醒") = vbCancel Then GoTo out
            End If
            If tmpW < 1 Then
                MsgBox "寬不得為0!!!", 48, "不能這樣"
                GoTo out
            End If
            Call Form14.CDrawer.NewFile(tmpW, tmpH, tmpZ, 5)
            Call Form14.SetUIPos
            Call Form14.EnableEdit(True)
            Call Form14.Refresh_ComboPage
            Call Form14.Command2_Click(3)
            Unload Me
        End If
        Form14.IsChanged = False
        Call Form14.SetMainCaption
    End If
    Exit Sub
out:
    Debug.Print "建立新檔 out"
End Sub

Private Sub Option1_Click(Index As Integer)
    If Option1(1).value = True Then
        Option3(2).value = True
        Option3(0).Enabled = False
        Option3(1).Enabled = False
        Option3(3).Enabled = False
    Else
        Option3(0).Enabled = True
        Option3(1).Enabled = True
        Option3(3).Enabled = True
    End If
    
End Sub
