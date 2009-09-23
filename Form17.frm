VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form17 
   BorderStyle     =   5  '可調整工具視窗
   Caption         =   "轉換圖片到畫布"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   285
   ClientWidth     =   4680
   LinkTopic       =   "Form17"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所屬視窗中央
   Begin VB.CheckBox Check1 
      Caption         =   "去背"
      Height          =   210
      Left            =   255
      TabIndex        =   13
      Top             =   2325
      Width           =   825
   End
   Begin VB.CommandButton Command1 
      Caption         =   "關閉"
      Height          =   315
      Index           =   1
      Left            =   2505
      TabIndex        =   12
      Top             =   2685
      Width           =   1425
   End
   Begin VB.CommandButton Command1 
      Caption         =   "進行轉換"
      Height          =   315
      Index           =   0
      Left            =   765
      TabIndex        =   8
      Top             =   2685
      Width           =   1425
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Windows轉換法"
      Height          =   195
      Index           =   1
      Left            =   330
      TabIndex        =   6
      Top             =   1950
      Width           =   1920
   End
   Begin VB.OptionButton Option1 
      Caption         =   "一般"
      Height          =   195
      Index           =   0
      Left            =   285
      TabIndex        =   3
      Top             =   480
      Value           =   -1  'True
      Width           =   720
   End
   Begin VB.Frame Frame1 
      Caption         =   "色彩判斷"
      Height          =   1380
      Left            =   165
      TabIndex        =   0
      Top             =   480
      Width           =   4395
      Begin VB.CheckBox Check2 
         Caption         =   "鎖定範圍大小"
         Height          =   210
         Left            =   780
         TabIndex        =   16
         Top             =   1050
         Width           =   1380
      End
      Begin VB.CommandButton Command2 
         Caption         =   "預設值"
         Height          =   270
         Left            =   3240
         TabIndex        =   9
         Top             =   1005
         Width           =   870
      End
      Begin MSComctlLib.Slider Slider1 
         Height          =   255
         Index           =   0
         Left            =   705
         TabIndex        =   1
         Top             =   285
         Width           =   2790
         _ExtentX        =   4921
         _ExtentY        =   450
         _Version        =   393216
         Max             =   255
         TickFrequency   =   5
      End
      Begin MSComctlLib.Slider Slider1 
         Height          =   255
         Index           =   1
         Left            =   705
         TabIndex        =   2
         Top             =   660
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   450
         _Version        =   393216
         Max             =   255
         TickFrequency   =   5
      End
      Begin VB.Label Label5 
         Height          =   210
         Left            =   2250
         TabIndex        =   17
         Top             =   1050
         Width           =   525
      End
      Begin VB.Label Label4 
         Caption         =   "Label4"
         Height          =   165
         Index           =   1
         Left            =   3495
         TabIndex        =   15
         Top             =   675
         Width           =   480
      End
      Begin VB.Label Label4 
         Caption         =   "Label4"
         Height          =   165
         Index           =   0
         Left            =   3495
         TabIndex        =   14
         Top             =   315
         Width           =   450
      End
      Begin VB.Label Label3 
         Height          =   180
         Index           =   1
         Left            =   4110
         TabIndex        =   11
         Top             =   690
         Width           =   210
      End
      Begin VB.Label Label3 
         Height          =   180
         Index           =   0
         Left            =   4110
         TabIndex        =   10
         Top             =   330
         Width           =   210
      End
      Begin VB.Label Label1 
         Caption         =   "亮限"
         Height          =   240
         Index           =   1
         Left            =   210
         TabIndex        =   5
         Top             =   690
         Width           =   405
      End
      Begin VB.Label Label1 
         Caption         =   "暗限"
         Height          =   240
         Index           =   0
         Left            =   210
         TabIndex        =   4
         Top             =   345
         Width           =   405
      End
   End
   Begin VB.Label Label2 
      Caption         =   "選擇色彩轉換方式"
      Height          =   225
      Left            =   135
      TabIndex        =   7
      Top             =   240
      Width           =   1860
   End
End
Attribute VB_Name = "Form17"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rangSize As Integer
Public Sub RefreshUI()
    Slider1(0).value = Form14.CDrawer.CA_v2_GetCaVars(6)
    Call Slider1_Scroll(0)
    Slider1(1).value = Form14.CDrawer.CA_v2_GetCaVars(7)
    Call Slider1_Scroll(1)
    'debug.Print Form14.CDrawer.CA_v2_GetCaVars(6)
End Sub



Private Sub Command1_Click(Index As Integer)
    Select Case Index
        Case 0
            If Option1(0).value = True Then
                Call Form14.CDrawer.CA_v2_SetCaVars(6, Slider1(0).value)
                Call Form14.CDrawer.CA_v2_SetCaVars(7, Slider1(1).value)
                Call Form14.CDrawer.Pic2VAC_V1(Check1.value)
            Else
                Call Form14.CDrawer.Pic2VAC_win(Check1.value)
            End If
        Case 1
            Unload Me
    End Select
End Sub

Private Sub Command2_Click()
    Call Form14.CDrawer.CA_v2_BC_DEFUALT
    Call RefreshUI
End Sub
Private Sub Form_Load()
    Call RefreshUI
End Sub

Private Sub Slider1_Scroll(Index As Integer)
    Label4(0).Caption = Slider1(0).value
    Label4(1).Caption = Slider1(1).value
    If Slider1(0).value >= Slider1(1).value Then Slider1(1 - Index).value = Slider1(Index).value
    Label3(0).BackColor = RGB(Slider1(0).value, Slider1(0).value, Slider1(0).value)
    Label3(1).BackColor = RGB(Slider1(1).value, Slider1(1).value, Slider1(1).value)
    If Check2.value = 1 Then
        If Index = 0 Then
            Slider1(1).value = Slider1(0).value + rangSize
        Else
            Slider1(0).value = Slider1(1).value - rangSize
        End If
    Else
        rangSize = Slider1(1).value - Slider1(0).value
        Label5.Caption = rangSize
    End If
End Sub
