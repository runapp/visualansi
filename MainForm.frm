VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.MDIForm MainForm 
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   4650
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   6705
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  '系統預設值
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  '對齊表單上方
      Height          =   630
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6705
      _ExtentX        =   11827
      _ExtentY        =   1111
      ButtonWidth     =   609
      ButtonHeight    =   953
      Appearance      =   1
      _Version        =   393216
      OLEDropMode     =   1
   End
   Begin VB.Menu Me_File 
      Caption         =   "檔案(&F)"
      Begin VB.Menu Me_Open 
         Caption         =   "開啟舊檔(&O)"
      End
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Canvas_set As New Collection
Public IndexCounter As Integer
Private Sub MDIForm_Load()
    IndexCounter = 0
End Sub

Private Sub Me_Open_Click()
    'Dim tmp As New frmCanvas
    'tmp.Show
    'Canvas_set.Add tmp, "1"
    'IndexCounter = IndexCounter + 1
    'Canvas_set.Item("1").Show
End Sub
