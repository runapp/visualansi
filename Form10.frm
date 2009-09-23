VERSION 5.00
Begin VB.Form Form10 
   BorderStyle     =   4  '單線固定工具視窗
   Caption         =   "選項"
   ClientHeight    =   3030
   ClientLeft      =   45
   ClientTop       =   270
   ClientWidth     =   3555
   LinkTopic       =   "Form10"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   3555
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton Command1 
      Caption         =   "取消"
      Height          =   315
      Index           =   1
      Left            =   1950
      TabIndex        =   7
      Top             =   2580
      Width           =   945
   End
   Begin VB.CommandButton Command1 
      Caption         =   "確定"
      Height          =   300
      Index           =   0
      Left            =   405
      TabIndex        =   6
      Top             =   2595
      Width           =   945
   End
   Begin VB.Frame Frame2 
      Caption         =   "習慣"
      Height          =   645
      Left            =   120
      TabIndex        =   3
      Top             =   1860
      Width           =   3375
      Begin VB.CheckBox Check1 
         Caption         =   "關閉時詢問儲存"
         Height          =   180
         Left            =   195
         TabIndex        =   4
         Top             =   300
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "顯示"
      Height          =   1680
      Left            =   105
      TabIndex        =   0
      Top             =   90
      Width           =   3360
      Begin VB.CheckBox Check3 
         Caption         =   "工具屬性置底(螢幕不夠大時使用)"
         Height          =   285
         Left            =   210
         TabIndex        =   9
         Top             =   1245
         Width           =   3075
      End
      Begin VB.CheckBox Check2 
         Caption         =   "不顯示選擇框線(不包括物件模式)"
         Height          =   390
         Left            =   210
         TabIndex        =   5
         Top             =   840
         Width           =   3030
      End
      Begin VB.ComboBox Combo1 
         Height          =   300
         Left            =   990
         Style           =   2  '單純下拉式
         TabIndex        =   2
         Top             =   225
         Width           =   570
      End
      Begin VB.Label Label2 
         Caption         =   "(下次啟動時這項設定才會生效)"
         Height          =   270
         Left            =   195
         TabIndex        =   8
         Top             =   585
         Width           =   2490
      End
      Begin VB.Label Label1 
         Caption         =   "字體大小"
         Height          =   255
         Left            =   225
         TabIndex        =   1
         Top             =   300
         Width           =   780
      End
   End
End
Attribute VB_Name = "Form10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click(Index As Integer)
If Index = 0 Then
    SysInfo.Frontsize = Combo1.text
    SysInfo.HideSelect = Check2.Value
    SysInfo.CheckSave = Check1.Value
    SysInfo.ToolPBoxDown = Check3.Value
End If
If SysInfo.HideSelect = 1 And SysInfo.EdMode < 6 Then
    Form1.Shape3.Visible = False
Else
    Form1.Shape3.Visible = True
End If
Unload Form10
End Sub

Private Sub Form_Load()
Combo1.AddItem 12
Combo1.AddItem 13
Combo1.AddItem 14
Combo1.AddItem 15
Combo1.ListIndex = 1
Debug.Print "SysInfo.Frontsize=" & SysInfo.Frontsize
Combo1.ListIndex = SysInfo.Frontsize - 12
Check2.Value = SysInfo.HideSelect
Check1.Value = SysInfo.CheckSave
 Check3.Value = SysInfo.ToolPBoxDown
End Sub

