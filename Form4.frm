VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form4 
   BorderStyle     =   4  '單線固定工具視窗
   Caption         =   "建立新檔案"
   ClientHeight    =   3390
   ClientLeft      =   45
   ClientTop       =   270
   ClientWidth     =   3855
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3390
   ScaleWidth      =   3855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所屬視窗中央
   Begin VB.CommandButton Command2 
      Caption         =   "取消"
      Height          =   315
      Left            =   2220
      TabIndex        =   12
      Top             =   2970
      Width           =   795
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4260
      Top             =   570
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form4.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form4.frx":08DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form4.frx":0BF8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command1 
      Caption         =   "確定"
      Height          =   330
      Left            =   555
      TabIndex        =   11
      Top             =   2955
      Width           =   870
   End
   Begin VB.Frame Frame1 
      Caption         =   "屬性"
      Height          =   1815
      Left            =   75
      TabIndex        =   1
      Top             =   1050
      Width           =   3690
      Begin VB.TextBox Text1 
         Alignment       =   2  '置中對齊
         Height          =   270
         Index           =   3
         Left            =   2670
         TabIndex        =   15
         Text            =   "78"
         Top             =   540
         Width           =   345
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  '置中對齊
         Height          =   270
         Index           =   2
         Left            =   3210
         TabIndex        =   14
         Text            =   "21"
         Top             =   540
         Width           =   360
      End
      Begin VB.OptionButton Option1 
         Caption         =   "全頁(80X23)"
         Height          =   225
         Index           =   3
         Left            =   150
         TabIndex        =   13
         Top             =   885
         Width           =   1320
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  '置中對齊
         Height          =   255
         Left            =   1770
         TabIndex        =   10
         Text            =   "2"
         Top             =   1305
         Width           =   525
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  '置中對齊
         Height          =   270
         Index           =   1
         Left            =   3000
         TabIndex        =   6
         Text            =   "14"
         Top             =   900
         Width           =   360
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  '置中對齊
         Height          =   270
         Index           =   0
         Left            =   2385
         TabIndex        =   5
         Text            =   "28"
         Top             =   900
         Width           =   345
      End
      Begin VB.OptionButton Option1 
         Caption         =   "自訂"
         Height          =   225
         Index           =   2
         Left            =   1710
         TabIndex        =   4
         Top             =   915
         Width           =   705
      End
      Begin VB.OptionButton Option1 
         Caption         =   "動畫全頁"
         Height          =   225
         Index           =   1
         Left            =   1695
         TabIndex        =   3
         Top             =   570
         Width           =   1080
      End
      Begin VB.OptionButton Option1 
         Caption         =   "簽名檔(80 X 6)"
         Height          =   225
         Index           =   0
         Left            =   150
         TabIndex        =   2
         Top             =   600
         Width           =   1530
      End
      Begin VB.Label Label4 
         Caption         =   "X"
         Height          =   255
         Left            =   3060
         TabIndex        =   16
         Top             =   570
         Width           =   210
      End
      Begin VB.Label Label3 
         Caption         =   "預設頁數"
         Height          =   195
         Left            =   990
         TabIndex        =   9
         Top             =   1350
         Width           =   810
      End
      Begin VB.Label Label2 
         Caption         =   "畫布大小(寬X高)"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   300
         Width           =   1620
      End
      Begin VB.Label Label1 
         Caption         =   "X"
         Height          =   180
         Left            =   2805
         TabIndex        =   7
         Top             =   945
         Width           =   180
      End
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   4350
      Top             =   1410
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form4.frx":0F14
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form4.frx":17F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form4.frx":1B0C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   885
      Left            =   75
      TabIndex        =   0
      Top             =   105
      Width           =   3705
      _ExtentX        =   6535
      _ExtentY        =   1561
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
    On Error GoTo out
    If Text1(2).text <> "21" And ListView1.SelectedItem.Index = 3 Then
        If MsgBox("除非你要輸出成ptt的pmore動畫，否則一般pagedown動畫請將行數由" & Text1(2).text & "改為21，你要依目前設定動畫行數嗎?", vbOKCancel, "警告") = vbCancel Then Exit Sub
    End If
    Call SendFInfo
    OFP.Closed = False
    Unload Form4
    Exit Sub
out:
    Debug.Print "建立新檔->確定 Error Out"
End Sub

Private Sub Command2_Click()
    Unload Form4
End Sub

Private Sub Form_Load()
    Set ListView1.Icons = ImageList1
    Set ListView1.SmallIcons = ImageList2
    ListView1.ListItems.Add , , "單頁畫", 2, 2
    ListView1.ListItems.Add , , "多頁畫", 3, 3
    ListView1.ListItems.Add , , "多頁動畫", 1, 1
    SetAnimC ListView1.SelectedItem.Index
    Option1(1).Value = True
End Sub

Private Sub ListView1_DblClick()
Debug.Print "ListView1_DblClick"
End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Debug.Print "ListView1_ItemClick" & Item.Index
    SetAnimC Item.Index
End Sub

Private Sub SendFInfo()
    Dim FInfo As VAFileInfo
    FInfo.IDC = 827
    Select Case ListView1.SelectedItem.Index
        Case Is = 1
            FInfo.filetype = 1
            FInfo.ArrZLenth = 1
        Case Is = 2
            FInfo.filetype = 3
            FInfo.ArrZLenth = Val(Text2.text)
        Case Is = 3
            FInfo.filetype = 2
            FInfo.ArrZLenth = Val(Text2.text)
    End Select
    
    If Option1(0).Value = True Then
        FInfo.ArrXUbound = 79
        FInfo.ArrYUbound = 5
    End If
    If Option1(1).Value = True Then
        FInfo.ArrXUbound = Val(Text1(3).text) - 1
        FInfo.ArrYUbound = Val(Text1(2).text) - 1
    End If
    
    If Option1(2).Value = True Then
        FInfo.ArrXUbound = Val(Text1(0).text) - 1
        FInfo.ArrYUbound = Val(Text1(1).text) - 1
    End If
    If Option1(3).Value = True Then
        FInfo.ArrXUbound = 79
        FInfo.ArrYUbound = 22
    End If
    
    CreatNewFile Arrf, FInfo

End Sub

Private Sub Option1_Click(Index As Integer)
If Option1(2).Value = True Then
    Text1(0).Enabled = True
    Text1(1).Enabled = True
    Label1.Enabled = True
Else
    Text1(0).Enabled = False
    Text1(1).Enabled = False
    Label1.Enabled = False

End If
End Sub

Public Sub SetAnimC(ByVal Index As Integer)

Select Case Index

    Case 1
        Label3.Enabled = False
        Text2.Enabled = False
        Option1(0).Enabled = True
        Option1(2).Enabled = True
        Option1(3).Enabled = True
    Case 2
        Label3.Enabled = True
        Text2.Enabled = True
        Option1(0).Enabled = True
        Option1(2).Enabled = True
        Option1(3).Enabled = True
    Case 3
        Label3.Enabled = True
        Text2.Enabled = True
        Option1(0).Enabled = False
        Option1(2).Enabled = False
        Option1(3).Enabled = False
        Option1(1).Value = True

End Select
End Sub
