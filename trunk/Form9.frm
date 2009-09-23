VERSION 5.00
Begin VB.Form Form9 
   BorderStyle     =   1  '單線固定
   Caption         =   "關於Visual Ansi"
   ClientHeight    =   4470
   ClientLeft      =   345
   ClientTop       =   390
   ClientWidth     =   3885
   Icon            =   "Form9.frx":0000
   LinkTopic       =   "Form9"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4470
   ScaleWidth      =   3885
   StartUpPosition =   2  '螢幕中央
   Begin VB.CommandButton Command1 
      Caption         =   "關閉"
      Height          =   300
      Left            =   1200
      Style           =   1  '圖片外觀
      TabIndex        =   0
      Top             =   4065
      Width           =   1410
   End
   Begin VB.Label Label1 
      Height          =   1875
      Left            =   255
      TabIndex        =   1
      Top             =   2055
      Width           =   3345
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  '單線固定
      Height          =   1905
      Left            =   165
      Picture         =   "Form9.frx":030A
      Stretch         =   -1  'True
      Top             =   60
      Width           =   3540
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Form9
End Sub

Private Sub Form_Load()
    Label1.Caption = "Visual Ansi 2008 apha ver. " & App.Major & "." & App.Minor & "." & App.Revision & " (2007/9/20)" _
    & vbCrLf & vbCrLf _
    & "Copyright (C) 2003-2007 Nerv.Studio 吱" & vbCrLf _
    & "Powered by Nerv.Style" & vbCrLf & vbCrLf _
    & "     本軟體雖受著作權法及國際公約法保護,凡未經授權任意複製散佈本電腦程式的部份或全部,絕對...........不會怎樣." & vbCrLf _
    & " 特別感謝: Suzanne aokman puzpuzpi ..."
End Sub

'0.9.8
'+雙色字支援
'+輸入模式
'+剪貼簿支援

'0.9.3
'Fix
'+ANSI編輯器
'+ANSI 列表內容
'+覆蓋繪圖 橡皮擦 加速
Private Sub Label1_Click()

End Sub
