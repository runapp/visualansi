VERSION 5.00
Begin VB.Form Form16 
   BorderStyle     =   5  '可調整工具視窗
   Caption         =   "檢視 / 調整圖片"
   ClientHeight    =   5145
   ClientLeft      =   165
   ClientTop       =   390
   ClientWidth     =   6315
   LinkTopic       =   "Form16"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   343
   ScaleMode       =   3  '像素
   ScaleWidth      =   421
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所屬視窗中央
   Begin VB.Frame Frame2 
      Caption         =   "工具箱"
      Height          =   3450
      Left            =   4110
      TabIndex        =   13
      Top             =   30
      Width           =   2130
      Begin VB.CommandButton Command2 
         Caption         =   "變形(需選取)"
         Height          =   300
         Index           =   6
         Left            =   300
         TabIndex        =   25
         Top             =   3030
         Width           =   1425
      End
      Begin VB.CommandButton Command2 
         Caption         =   "設定位置"
         Height          =   300
         Index           =   5
         Left            =   300
         TabIndex        =   23
         Top             =   2625
         Width           =   1425
      End
      Begin VB.CommandButton Command2 
         Caption         =   "還原圖片"
         Height          =   300
         Index           =   4
         Left            =   285
         TabIndex        =   22
         Top             =   765
         Width           =   1425
      End
      Begin VB.CommandButton Command2 
         Caption         =   "設定左上角"
         Height          =   300
         Index           =   3
         Left            =   285
         TabIndex        =   21
         Top             =   2250
         Width           =   1425
      End
      Begin VB.CommandButton Command2 
         Caption         =   "剪裁(需選取)"
         Height          =   300
         Index           =   2
         Left            =   285
         TabIndex        =   20
         Top             =   1890
         Width           =   1425
      End
      Begin VB.CommandButton Command2 
         Caption         =   "最大方格"
         Height          =   300
         Index           =   1
         Left            =   285
         TabIndex        =   17
         Top             =   1515
         Width           =   1410
      End
      Begin VB.CommandButton Command2 
         Caption         =   "等比例最大"
         Height          =   300
         Index           =   0
         Left            =   285
         TabIndex        =   16
         Top             =   1155
         Width           =   1440
      End
      Begin VB.Label Label5 
         Height          =   195
         Left            =   135
         TabIndex        =   24
         Top             =   465
         Width           =   1950
      End
      Begin VB.Label Label4 
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   225
         Width           =   1935
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "手動修改"
      Height          =   1650
      Left            =   60
      TabIndex        =   1
      Top             =   3405
      Width           =   5265
      Begin VB.CommandButton Command1 
         Caption         =   "取消"
         Height          =   285
         Index           =   1
         Left            =   315
         TabIndex        =   18
         Top             =   1290
         Width           =   1290
      End
      Begin VB.CommandButton Command1 
         Caption         =   "套用"
         Height          =   285
         Index           =   0
         Left            =   315
         TabIndex        =   14
         Top             =   930
         Width           =   1290
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  '靠右對齊
         Height          =   270
         Index           =   3
         Left            =   4065
         TabIndex        =   5
         Top             =   855
         Width           =   555
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  '靠右對齊
         Height          =   270
         Index           =   2
         Left            =   3270
         TabIndex        =   4
         Top             =   840
         Width           =   555
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  '靠右對齊
         Height          =   270
         Index           =   1
         Left            =   1125
         TabIndex        =   3
         Text            =   "0"
         Top             =   555
         Width           =   585
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  '靠右對齊
         Height          =   270
         Index           =   0
         Left            =   285
         TabIndex        =   2
         Text            =   "0"
         Top             =   540
         Width           =   600
      End
      Begin VB.Label Label1 
         Caption         =   "來源圖: "
         Height          =   225
         Index           =   5
         Left            =   2025
         TabIndex        =   15
         Top             =   915
         Width           =   1125
      End
      Begin VB.Label Label3 
         Caption         =   "X:               Y:"
         Height          =   195
         Left            =   90
         TabIndex        =   12
         Top             =   585
         Width           =   1185
      End
      Begin VB.Label Label2 
         Caption         =   "縮放率"
         Height          =   195
         Left            =   3300
         TabIndex        =   11
         Top             =   1215
         Width           =   1920
      End
      Begin VB.Label Label1 
         Caption         =   "X"
         Height          =   225
         Index           =   4
         Left            =   3870
         TabIndex        =   10
         Top             =   885
         Width           =   135
      End
      Begin VB.Label Label1 
         Caption         =   "目前: "
         Height          =   225
         Index           =   3
         Left            =   3255
         TabIndex        =   9
         Top             =   570
         Width           =   1620
      End
      Begin VB.Label Label1 
         Caption         =   "來源圖: "
         Height          =   225
         Index           =   2
         Left            =   2025
         TabIndex        =   8
         Top             =   585
         Width           =   810
      End
      Begin VB.Label Label1 
         Caption         =   "寬X高(單位 : 像素)"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   1980
         TabIndex        =   7
         Top             =   255
         Width           =   1740
      End
      Begin VB.Label Label1 
         Caption         =   "位置(單位 : 像素)"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   60
         TabIndex        =   6
         Top             =   255
         Width           =   1905
      End
   End
   Begin VB.PictureBox vPic 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   3360
      Left            =   0
      MousePointer    =   2  '十字形狀
      ScaleHeight     =   220
      ScaleMode       =   3  '像素
      ScaleWidth      =   268
      TabIndex        =   0
      Top             =   0
      Width           =   4080
      Begin VB.Shape Shape3 
         BorderColor     =   &H00FF00FF&
         BorderStyle     =   3  '點線
         DrawMode        =   4  'Mask Not Pen
         Height          =   150
         Left            =   150
         Top             =   240
         Width           =   150
      End
   End
   Begin VB.Menu Me_LoadPic 
      Caption         =   "載入圖片"
   End
   Begin VB.Menu Me_BG 
      Caption         =   "設為背景"
   End
   Begin VB.Menu ME_SavePic 
      Caption         =   "另存圖片"
   End
   Begin VB.Menu Me_Close 
      Caption         =   "關閉"
   End
End
Attribute VB_Name = "Form16"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'隱藏關閉鈕
Private Type MENUITEMINFO
  cbSize As Long
  fMask As Long
  fType As Long
  fState As Long
  wID As Long
  hSubMenu As Long
  hbmpChecked As Long
  hbmpUnchecked As Long
  dwItemData As Long
  dwTypeData As String
  cch As Long
End Type
Private Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Private Declare Function GetMenuItemInfo Lib "user32" Alias "GetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal B As Boolean, lpmii As MENUITEMINFO) As Long
Private Declare Function SetMenuItemInfo Lib "user32" Alias "SetMenuItemInfoA" (ByVal hMenu As Long, ByVal uItem As Long, ByVal fByPosition As Long, lpmii As MENUITEMINFO) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const WM_NCACTIVATE = &H86
Private Const SC_CLOSE = &HF060&
Private Const MIIM_STATE = &H1&
Private Const MIIM_ID = &H2&
Private Const MFS_GRAYED = &H3&
Private Const MFS_CHECKED = &H8&
Private Const xMenuID = 10&
'Private Const MIIM_STATE = &H1
'Private Const SC_CLOSE = &HF060
Private SL As New SelectLine
Dim IPED_TAG As Byte
Dim CDrawer As CubeDrawerObject
Dim MousePos(1) As Long '紀錄滑鼠的位置(為了轉成整數)

Private Sub Command1_Click(Index As Integer)
    Select Case Index
        Case 0
            Dim tmpVars(3) As Long
            tmpVars(0) = Text1(0).text
            tmpVars(1) = Text1(1).text
            tmpVars(2) = Text1(2).text
            tmpVars(3) = Text1(3).text
            Call CDrawer.IPED_SetVars(tmpVars(0), tmpVars(1), tmpVars(2), tmpVars(3))
        Case 1
            Call PicInfo_Refresh
    End Select
End Sub

Private Sub Command2_Click(Index As Integer)

    Select Case Index
        Case 0
            Call CDrawer.IPED_FitPBSize
        Case 1
            Call CDrawer.IPED_SizePixelCube
            'Call CDrawer.Pic2VAC_win
        Case 2
            '剪裁
            Dim tmpW As Long, tmpH As Long  '
            Call SL.CorrectPoints
            tmpW = SL.EndPoint_X - SL.StartPoint_X + 1
            tmpH = SL.EndPoint_Y - SL.StartPoint_Y + 1
            If tmpW = 1 Or tmpH = 1 Then
                MsgBox "請先選取範圍", 48, "無法剪裁"
            Else
               Call CDrawer.IPED_Cut(SL.StartPoint_X, SL.StartPoint_Y, tmpW, tmpH)
            End If
        Case 3
            '設定左上角頂點
            If IPED_TAG = 0 Then
                IPED_TAG = 3
                Command2(3).Caption = "結束設定"
                For i = 0 To 5
                    If i <> 3 Then Command2(i).Enabled = False
                Next i
                Frame1.Enabled = False
            Else
                IPED_TAG = 0
                Command2(3).Caption = "設定左上角"
                For i = 0 To 5
                    Command2(i).Enabled = True
                Next i
                Frame1.Enabled = True
            End If
        Case 4
            '還原
            Call CDrawer.IPED_Restore
        Case 5
            If IPED_TAG = 0 Then
                IPED_TAG = 5
                Command2(5).Caption = "結束設定"
                For i = 0 To 4
                     Command2(i).Enabled = False
                Next i
                Frame1.Enabled = False
            Else
                IPED_TAG = 0
                Command2(5).Caption = "設定位置"
                For i = 0 To 5
                    Command2(i).Enabled = True
                Next i
                Frame1.Enabled = True
            End If
        Case 6
            Call SL.CorrectPoints
            If SL.StartPoint_X = SL.EndPoint_X And SL.StartPoint_Y = SL.EndPoint_Y Then
                MsgBox "請先在畫布選取一個範圍", 48, "無法變形"
            Else
                Call CDrawer.IPED_TransForm(SL.StartPoint_X, SL.StartPoint_Y, SL.EndPoint_X, SL.EndPoint_Y)
            End If
    End Select
    Call PicInfo_Refresh
End Sub

Private Sub Form_Load()
    '隱藏關閉鈕
    Dim hMenu As Long, MII As MENUITEMINFO
    hMenu = GetSystemMenu(Me.hwnd, 0)
    MII.cbSize = Len(MII)
    MII.dwTypeData = String(80, 0)
    MII.cch = Len(MII.dwTypeData)
    MII.fMask = MIIM_STATE
    MII.wID = SC_CLOSE
    
    GetMenuItemInfo hMenu, SC_CLOSE, False, MII
    
    MII.wID = xMenuID
    MII.fMask = MIIM_ID
    SetMenuItemInfo hMenu, SC_CLOSE, False, MII
    
    MII.fState = MII.fState Or MFS_GRAYED
    MII.fMask = MIIM_STATE
    SetMenuItemInfo hMenu, MII.wID, False, MII
    
    SendMessage Me.hwnd, WM_NCACTIVATE, True, ByVal 0&
    
    Set CDrawer = Form14.CDrawer
    Call CDrawer.Load_IP_PB(vPic)
    Me.ScaleMode = Form14.ScaleMode
    'Call ReSize_UI
    'Call ShowPic
    
    SL.TragetShape = Shape3
End Sub

Private Sub Picture1_Click()
    
End Sub

Public Sub ShowPic()
    CDrawer.ShowIPout2PB
    Call PicInfo_Refresh
    vPic.Refresh
End Sub

Public Sub ReSize_UI()
    vPic.Width = Form14.Picture1.Width
    vPic.Height = Form14.Picture1.Height
    Frame2.Left = vPic.Width
    If vPic.Height < Frame2.Top + Frame2.Height Then
        Frame1.Top = Frame2.Top + Frame2.Height
    Else
        Frame1.Top = vPic.Height
    End If
    Form16.Width = (Frame2.Left + Frame2.Width) * Screen.TwipsPerPixelX + 100
    'Debug.Print "W:" & (Frame1.Left + Frame1.Width) & " H:" & vPic.Height
    Form16.Height = (Frame1.Top + Frame1.Height) * Screen.TwipsPerPixelY + 800
    'Debug.Print "FW:" & Me.Width & " FH:" & Me.Height
End Sub


Public Sub PicInfo_Refresh()
On Error GoTo out
    Dim tmpSi(3) As Single
    Text1(0) = CDrawer.GetIP_Vars(0)
    Text1(1) = CDrawer.GetIP_Vars(1)
    Text1(2) = CDrawer.GetIP_Vars(2)
    tmpSi(0) = CDrawer.GetIP_Vars(2)
    Text1(3) = CDrawer.GetIP_Vars(3)
    tmpSi(1) = CDrawer.GetIP_Vars(3)
    Label1(5).Caption = CDrawer.GetIP_Vars(4) & "X" & CDrawer.GetIP_Vars(5)
    tmpSi(2) = CDrawer.GetIP_Vars(4)
    tmpSi(3) = CDrawer.GetIP_Vars(5)
    Label2.Caption = "縮放率:" & tmpSi(0) / tmpSi(2) & "X" & tmpSi(1) / tmpSi(3)
Exit Sub
out:
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Form16.Hide
End Sub

Private Sub Me_BG_Click()
    Call Form14.Command12_Click(1)
End Sub

Private Sub Me_Close_Click()
    Me.Hide
End Sub

Private Sub Me_LoadPic_Click()
    Call Form14.Command11_Click
End Sub

Private Sub ME_SavePic_Click()
On Error GoTo out

    Form14.CDialog1.DialogTitle = "另存圖片"
    Form14.CDialog1.Filter = "*.bmp(點陣圖檔)|*.bmp"
    Form14.CDialog1.FileName = ""
    Form14.CDialog1.ShowSave
    If Dir(Form14.CDialog1.FileName, vbHidden Or vbDirectory Or vbReadOnly Or vbSystem) <> "" Then
        If MsgBox("檔案已經存在 請問你要覆蓋他嗎?", 49, "檔案已存在") = vbCancel Then Exit Sub
    End If
    Debug.Print Form14.CDialog1.FileName
    vPic.AutoRedraw = True
    
    'vPic.PaintPicture
    'vPic.Picture = vPic.Image
    SavePicture vPic.Image, Form14.CDialog1.FileName
    'vPic.AutoRedraw = True
    Exit Sub
out:
    Debug.Print "ME_SavePic_Click OUT "
    vPic.AutoRedraw = True
    'Form14.CDialog1
End Sub

Private Sub vPic_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MousePos(0) = X
    MousePos(1) = Y
    'Shape3.Visible = False
    SL.EndPoint_Y = MousePos(1)
    SL.EndPoint_X = MousePos(0)
    SL.StartPoint_X = MousePos(0)
    SL.StartPoint_Y = MousePos(1)

    Call SL.DrawSelect
    Label5.Caption = "(" & SL.StartPoint_X & "," & SL.StartPoint_Y & ")-(" & SL.EndPoint_X & "," & SL.EndPoint_Y & ")"
    If IPED_TAG = 3 Then
        Call CDrawer.IPED_SetTopLeft(MousePos(0), MousePos(1))
    End If
    If IPED_TAG = 5 Then
        Call CDrawer.IPED_SetPos(MousePos(0), MousePos(1))
        
    End If
End Sub

Private Sub vPic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MousePos(0) = X
    MousePos(1) = Y
    Label4.Caption = "(" & MousePos(0) & "," & MousePos(1) & ")"
    If Button = 1 Then
        SL.EndPoint_X = MousePos(0)
        SL.EndPoint_Y = MousePos(1)
        SL.DrawSelect
        Label5.Caption = "(" & SL.StartPoint_X & "," & SL.StartPoint_Y & ")-(" & SL.EndPoint_X & "," & SL.EndPoint_Y & ")"
    End If
End Sub
