VERSION 5.00
Begin VB.Form frmCanvas 
   BackColor       =   &H8000000A&
   Caption         =   "Visual Ansi 0.9.7a"
   ClientHeight    =   4020
   ClientLeft      =   165
   ClientTop       =   165
   ClientWidth     =   4470
   Icon            =   "frmCanvas.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4020
   ScaleWidth      =   4470
   Tag             =   "0"
   Begin VB.Frame Frame6 
      Caption         =   "BBS"
      Height          =   1875
      Left            =   570
      TabIndex        =   3
      Top             =   6285
      Visible         =   0   'False
      Width           =   6675
      Begin VB.CommandButton Command4 
         Caption         =   "複製"
         Height          =   270
         Left            =   1860
         TabIndex        =   6
         Top             =   300
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "轉成BBSAnsi"
         Height          =   270
         Left            =   330
         TabIndex        =   5
         Top             =   300
         Width           =   1410
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  '平面
         Height          =   960
         Left            =   330
         MultiLine       =   -1  'True
         ScrollBars      =   3  '兩者皆有
         TabIndex        =   4
         Top             =   660
         Width           =   5910
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "畫布"
      Height          =   3465
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   3300
      Begin VB.PictureBox Pic1 
         Appearance      =   0  '平面
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         FillColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "細明體"
            Size            =   14.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   2820
         Left            =   90
         ScaleHeight     =   186
         ScaleMode       =   3  '像素
         ScaleWidth      =   202
         TabIndex        =   1
         Top             =   435
         Width           =   3060
         Begin VB.Shape Shape3 
            BorderColor     =   &H00FF00FF&
            BorderStyle     =   3  '點線
            DrawMode        =   4  'Mask Not Pen
            Height          =   315
            Left            =   0
            Top             =   0
            Width           =   135
         End
      End
      Begin VB.Label Label2 
         Height          =   210
         Left            =   210
         TabIndex        =   2
         Top             =   210
         Width           =   720
      End
   End
End
Attribute VB_Name = "frmCanvas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim newf
Dim ForColor As Byte
Dim BacColor As Byte
Dim Pic1MouseDown As Boolean
Dim AnsiFile As Integer
Dim OpFile As Integer
Dim FileOpMode As Integer
Dim NowAnsi As String
Dim ToolP As Frame
Public SL As New SelectLine
Dim AD As New ApiDrawObject
Public FileSys As New FileSystemObject
Dim CPArr() As ColorLayer   '複製貼上所用之記憶體



Private Sub Form_Load()

    'For i = 0 To 9
        ' MSFlexGrid1.ColWidth(i) = 255
    'Next i
    'Call LoadAnsi(App.Path & "\Ansi.txt", Me.MSFlexGrid1)

'=====預設值的設定====
    SysInfo.ForColor = 7
    SysInfo.Frontsize = 14
    SysInfo.EdMode = 1 '設定工具
    'Set ToolP = Frame3 '設定工具屬性欄
'    Toolbar1.Buttons(1).value = tbrPressed
'    Toolbar2.Buttons(1).value = tbrPressed
    '交換顏色的設定
    SysInfo.ExChColor.Color(0) = 7
    SysInfo.ExChColor.Color(1) = 7
    SysInfo.ExChColor.Color(2) = 7
    SysInfo.ExChColor.Color(3) = 7
    '前景來源-多行文字
    FString.StrLen(1) = 1
'=====================

    Call GetConfic
    Call SetToolP
    '設定調色盤
    Call SetColorBoard
    Call SetNowAnsi
    'SysInfo.PPA = 285
    Call ForScreenSize(SysInfo.Frontsize)
    
    Call SetSize(28, 14, 1, 1)
    '初始化物件暫存
    ReDim ObjCA(0, 0)
    '初始化剪貼簿
    ReDim CPArr(0, 0)
    'Command18(2).Tag = "0"
    '設定繪圖物件的目標
    AD.Traget = Pic1
    '設定區塊選擇物件
    SL.TragetShape = Shape3
    If SysInfo.HideSelect = 1 Then Shape3.Visible = False
    '讀取物件清單檔
    Call ObjList_Read
    '關閉狀態設置
    OFP.Closed = True
    'CheckClose
    Unload Form8
    Me.Show
    If Command <> "" Then
        Call OpenFile_Command(Command)
        'MsgBox "got command", 64, "information"
    End If
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If SysInfo.CheckSave = 1 And OFP.Closed = False Then
        Call AskSave
    
    End If
    Unload Form7    '卸除匯入文章
    Unload Form6    '卸除輸出
    Unload me3   '卸除文句內容設定
    '儲存物件清單檔
    Call ObjList_Save
    Call SetConfic
End Sub

Private Sub Label13_Click(Index As Integer)

End Sub





Private Sub Frame3_DragDrop(Source As Control, X As Single, Y As Single)
Debug.Print "Drog"
End Sub

Private Sub Frame1_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub Frame5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Pic1MouseDown = False
End Sub

Private Sub Lb_ExCh_Color_Click(Index As Integer)
    Lb_ExCh_Color(SysInfo.ExChColor.CurrentSel).BorderStyle = 0    '將前一個改成未選取
    SysInfo.ExChColor.CurrentSel = Index                    '設定要選取顏色的方塊
    Lb_ExCh_Color(SysInfo.ExChColor.CurrentSel).BorderStyle = 1    '將目前的方塊改成選取狀態
End Sub

Private Sub List1_Click()
    On Error GoTo out
    'If SysInfo.edmode = 7 Then
        Obj_ReadFile ObjList(List1.ListIndex).filepath, ObjCA
        'Debug.Print ObjList(List1.ListIndex).FilePath
    
    'End If
        Label8.Caption = UBound(ObjCA, 1) + 1 & "X" & UBound(ObjCA, 2) + 1
    Exit Sub
out:
    Debug.Print "List1_Click Error Out"
End Sub

Private Sub Me_About_Click()
    Form9.Show vbModal
End Sub

Private Sub Me_AnsiListEd_Click()
    Form3.Show vbModal
End Sub

Private Sub Me_Compile_Click()
    Form5.Show vbModal
End Sub

Private Sub Me_Director_Click()
    me1.Show vbModal
End Sub

Private Sub Me_Display_Click()
'撥放動畫
    Call DisplayVAA
End Sub

Private Sub Me_ImprortText_Click()
    Form7.Show
End Sub

Private Sub Me_New_Click()
    If OFP.Closed = False Then Call AskSave
    Call OpenNewFile
    'close 由開新檔案的部分判斷 防止取消的bug
    'OFP.Closed = False
    Call CheckClose
    Call Set_FileType_Visual
    
    Call Set_VAA_Combo
    Call VAA_SetButton
    Call ReDraw
    OFP.filepath = ""
    OFP.IsChanged = True
    Call SetFormCaption
    AD.SetTraget
    
End Sub

Private Sub Me_OpenFile_Click()
    On Error GoTo out
    If OFP.Closed = False Then Call AskSave
    FileOpMode = 1
    
    CDialog1.DialogTitle = "開啟舊檔"
    CDialog1.InitDir = App.Path
    CDialog1.Filter = "*.VAF(單頁畫) *.VAA(動畫) *.VAM(多頁畫)|*.vaf; *.vaa; *.vam|*.VAF(單頁畫)|*.vaf| *.VAA(動畫)| *.vaa|*.VAM(多頁畫)|*.vam"
    CDialog1.FileName = ""
    CDialog1.ShowOpen
    If FileSys.FileExists(CDialog1.FileName) = False Then
        MsgBox "請選擇存在的檔案", vbOKOnly, "檔案不存在"
        Exit Sub
    End If
    OFP.filepath = CDialog1.FileName
    Call SetFormCaption("載入中...")
    VA_ReadFile OFP.filepath, Arrf
    OFP.Closed = False
    OFP.IsChanged = False
    Call SetFormCaption
    Call CheckClose
    Call Set_FileType_Visual
    Call Set_VAA_Combo
    Call VAA_SetButton
    Call ReDraw
    
    Exit Sub
out:
    Debug.Print "Me_OpenFile_Click Error Out"

End Sub
Private Sub OpenFile_Command(ByVal CommandString As String)
On Error GoTo out
    If OFP.Closed = False Then Call AskSave
    FileOpMode = 1
    If FileSys.FileExists(CommandString) = False Then
        MsgBox "請選擇存在的檔案", vbOKOnly, "檔案不存在"
        Exit Sub
    End If
    OFP.filepath = CommandString
    VA_ReadFile OFP.filepath, Arrf
    OFP.Closed = False
    OFP.IsChanged = False
    Call SetFormCaption
    Call CheckClose
    Call Set_FileType_Visual
    Call Set_VAA_Combo
    Call VAA_SetButton
    Call ReDraw
    
Exit Sub
out:
    Debug.Print "OpenFile_Command Error Out"

End Sub

Private Sub Me_Refresh_Click()
    Call ReDraw
End Sub

Private Sub Me_Save_Click()
On Error GoTo out
    CDialog1.DialogTitle = "儲存檔案"
    Select Case OFP.filetype
        Case Is = 1
            CDialog1.Filter = "*.VAF(單頁畫)|*.vaf"
        Case Is = 2
            CDialog1.Filter = "*.VAA(動畫)|*.vaa"
        Case Is = 3
            CDialog1.Filter = "*.VAM(多頁畫)|*.vam"
    End Select
    If OFP.filepath = "" Then
        CDialog1.FileName = ""
        CDialog1.ShowSave
        OFP.filepath = CDialog1.FileName
    End If
    VA_SaveFile OFP.filepath, Arrf, OFP
    OFP.IsChanged = False
    Call SetFormCaption
Exit Sub
out:
    Debug.Print "Me_Save Error Out"
End Sub

Private Sub Me_SaveAs_Click()
On Error GoTo out

CDialog1.DialogTitle = "另存新檔"
Select Case OFP.filetype
    Case Is = 1
        CDialog1.Filter = "*.VAF(單頁畫)|*.vaf"
    Case Is = 2
        CDialog1.Filter = "*.VAA(動畫)|*.vaa"
    Case Is = 3
        CDialog1.Filter = "*.VAM(多頁畫)|*.vaa"
End Select
CDialog1.FileName = ""
CDialog1.ShowSave



If FileSys.FileExists(CDialog1.FileName) = True Then
    If MsgBox("這個檔案已經存在,你確定要覆蓋它嗎?", vbOKCancel, "檔案已存在") = vbNo Then Exit Sub
End If
VA_SaveFile CDialog1.FileName, Arrf, OFP

Exit Sub
out:
Debug.Print "Me_SaveAs Error Out"
End Sub

Private Sub Me_SetOptions_Click()
me0.Show vbModal
End Sub

Private Sub MSFlexGrid1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Debug.Print MSFlexGrid1.MouseRow
'Debug.Print MSFlexGrid1.MouseCol
'Debug.Print MSFlexGrid1.RowSel
'Debug.Print MSFlexGrid1.ColSel
Text1.text = MSFlexGrid1.text
Label4.Caption = Tlen(Text1.text)
End Sub

Private Sub MSFlexGrid1_RowColChange()
Text1.text = MSFlexGrid1.text
Label4.Caption = Tlen(Text1.text)
End Sub

Private Sub Option1_Click()
Call SetNowAnsi
End Sub

Private Sub Option2_Click()
Call SetNowAnsi
End Sub

Private Sub Option3_Click()
Call SetNowAnsi
End Sub

Private Sub Option4_Click(Index As Integer)
    Command17(0).Enabled = Option4(1).value
End Sub

Private Sub Option5_Click(Index As Integer)
    Command17(1).Enabled = Option5(1).value
End Sub

Private Sub Option6_Click(Index As Integer)
    Command17(2).Enabled = Option6(1).value
End Sub

Private Sub Me_File_Click()

End Sub

Private Sub Pic1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo out
Dim intX As Integer
Dim intY As Integer

intX = Fix(X)
intY = Fix(Y)


Select Case SysInfo.EdMode
    Case 1
        If Check2.value = 1 Then Call DoDrawBC(intX, intY)
        If Check1.value = 1 Then
            If Option2.value = False Then
                Call DoDraw(intX, intY, NowAnsi)
            Else
                Call DoMutiDraw(intX, intY, FString.str, 0)
                Pic1.Refresh
            End If
        End If

    Case 2
        If Option4(0).value = True Then
            If Check2.value = 1 Then Call DoEreaseB(intX, intY)
            If Check1.value = 1 Then Call DoErease(intX, intY)
            Pic1.Refresh
        Else
            SL.StartPoint_X = intX
            SL.StartPoint_Y = intY
            SL.EndPoint_X = intX
            SL.EndPoint_Y = intY
            SL.DrawSelect
        End If
    Case 3
        If Option2.value <> True Then
        
            If Check1.value = 1 Then
                If Tlen(NowAnsi) = 2 Then
                    Call DoErease(intX + 1, intY)
                End If
                    Call DoErease(intX, intY)

                Call DoDraw(intX, intY, NowAnsi)
            End If
            
            If Check2.value = 1 Then
                If Tlen(NowAnsi) = 2 Then
                    Call DoEreaseB(intX + 1, intY)
                End If
                
                    Call DoEreaseB(intX, intY)
                
                Call DoDrawBC(intX, intY)
            End If
        Else
             Call DoMutiDraw(intX, intY, FString.str, 1)
        End If

        Pic1.Refresh
    Case 4
        If Check1.value = 1 Then Call DoDraw(intX, intY, NowAnsi)
        If Check2.value = 1 Then Call DoDrawBC(intX, intY)
        Call SelectAnsi(intX, intY)
    Case 5
        If Option5(0).value = True Then
            Call PaintColor(intX, intY)

        Else
            SL.StartPoint_X = intX
            SL.StartPoint_Y = intY
            SL.EndPoint_X = intX
            SL.EndPoint_Y = intY
            SL.DrawSelect
        End If
    Case 6
        SL.StartPoint_X = intX
        SL.StartPoint_Y = intY
        SL.EndPoint_X = intX
        SL.EndPoint_Y = intY
        SL.DrawSelect
    Case 7
        If Check4.value = 1 Then
        '去背
            Call CLArrayPaste_C(ObjCA(), Arrf(), intX, intY, OFP.CurrentPage)

        Else
            Call ObjLibPo(ObjCA(), Arrf(), intX, intY)
        End If
        Call ReDraw_Area(intX, intY, UBound(ObjCA, 1) + intX, UBound(ObjCA, 2) + intY)
    Case 8
        If Option6(0).value = True Then
            Call ExChColor_Draw(intX, intY)
        Else
            SL.StartPoint_X = intX
            SL.StartPoint_Y = intY
            SL.EndPoint_X = intX
            SL.EndPoint_Y = intY
            SL.DrawSelect
        End If
    Case 9 '複製&貼上
        If Command18(2).Tag = "0" Then
            SL.StartPoint_X = intX
            SL.StartPoint_Y = intY
            SL.EndPoint_X = intX
            SL.EndPoint_Y = intY
            SL.DrawSelect
        Else
            If Check5.value = 0 Then
                Call ObjLibPo(CPArr(), Arrf(), intX, intY)
            Else
                '去背模式
                Call CLArrayPaste_C(CPArr(), Arrf(), intX, intY, OFP.CurrentPage)
            
            End If
            Call ReDraw_Area(intX, intY, UBound(CPArr, 1) + intX, UBound(CPArr, 2) + intY)
        End If
End Select
'If SysInfo.EdMode <> 7 And SysInfo.EdMode <> 6 Then
'
'    Call SelectAnsi(intX, intY)
'End If
If SysInfo.EdMode <> 6 And Check1.value <> 0 And Check2.value <> 0 Then
    OFP.IsChanged = True
    Call SetFormCaption
End If
Pic1MouseDown = True
Pic1.Refresh
Exit Sub
out:
Debug.Print "Pic1_MouseDown Error Out"
End Sub

Private Sub Pic1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo out
    Label2.Caption = "(" & Fix(X) & "," & Fix(Y) & ")"
    Dim intX As Integer
    Dim intY As Integer
    
    intX = Fix(X)
    intY = Fix(Y)
    
    If Pic1MouseDown Then
        Select Case SysInfo.EdMode
            Case 2
                If Option4(0).value = True Then
                    If Check2.value = 1 Then Call DoEreaseB(intX, intY)
                    If Check1.value = 1 Then Call DoErease(intX, intY)
                    Pic1.Refresh
                Else
                    SL.EndPoint_X = intX
                    SL.EndPoint_Y = intY
                    SL.DrawSelect
                End If
            Case 4
                If Check1.value = 1 Then Call DoDraw(intX, intY, NowAnsi)
                If Check2.value = 1 Then Call DoDrawBC(intX, intY)
    
            Case 5
                If Option5(0).value = True Then
                    Call PaintColor(intX, intY)
                Else
                    SL.EndPoint_X = intX
                    SL.EndPoint_Y = intY
                    SL.DrawSelect
                End If
            Case 6
                SL.EndPoint_X = intX
                SL.EndPoint_Y = intY
                SL.DrawSelect
            Case 8
            
                If Option6(0).value = True Then
                    Call ExChColor_Draw(intX, intY)
                Else
                    SL.EndPoint_X = intX
                    SL.EndPoint_Y = intY
                    SL.DrawSelect
                End If
            Case 9
                If Command18(2).Tag = "0" Then
                    SL.EndPoint_X = intX
                    SL.EndPoint_Y = intY
                    SL.DrawSelect
                End If
        End Select
    
        Pic1.Refresh
    End If
    If (SysInfo.EdMode = 2 And Option4(0).value = True) Or (SysInfo.EdMode = 5 And Option5(0).value = True) Or (SysInfo.EdMode = 8 And Option6(0).value = True) Then
            Call SelectAnsi(intX, intY)
    End If
    
    If SysInfo.EdMode = 7 Then
        SL.StartPoint_X = intX
        SL.StartPoint_Y = intY
        SL.EndPoint_X = intX + UBound(ObjCA, 1)
        SL.EndPoint_Y = intY + UBound(ObjCA, 2)
        SL.DrawSelect
    End If
    If SysInfo.EdMode = 9 And Command18(2).Tag = "1" Then
        SL.StartPoint_X = intX
        SL.StartPoint_Y = intY
        SL.EndPoint_X = intX + UBound(CPArr, 1)
        SL.EndPoint_Y = intY + UBound(CPArr, 2)
        SL.DrawSelect
    End If
    If SysInfo.EdMode <> 7 And SysInfo.EdMode <> 6 And SysInfo.EdMode <> 9 And SysInfo.EdMode <> 2 And SysInfo.EdMode <> 5 And SysInfo.EdMode <> 8 Then
        SL.StartPoint_X = intX
        SL.StartPoint_Y = intY
        
        If Option2.value = True Then
            SL.EndPoint_Y = intY + FString.StrLen(1) - 1
            SL.EndPoint_X = intX + FString.StrLen(0) - 1
        Else
            SL.EndPoint_Y = intY
            SL.EndPoint_X = intX + Tlen(NowAnsi) - 1
        End If
        
        SL.DrawSelect
    End If
Exit Sub
out:
Debug.Print "Pic1_MouseMove Error Out"

End Sub

Public Sub lentest()
For i = 8 To 2000
    If Pic1.Point(i, 150) = 0 Then
        Debug.Print i
        Exit For
    End If

DoEvents
Next i
End Sub



Private Sub Pic1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Pic1MouseDown = False
If Check3.value = 1 Then
    Debug.Print "Arrf(" & Fix(X) & "," & Fix(Y) & ").Ansi=" & Arrf(Fix(X), Fix(Y), OFP.CurrentPage).Ansi
    Debug.Print "Arrf(" & Fix(X) & "," & Fix(Y) & ").Color=" & Arrf(Fix(X), Fix(Y), OFP.CurrentPage).Color
    Debug.Print QBCToAnsiC(Arrf(Fix(X), Fix(Y), OFP.CurrentPage).Color)
    Debug.Print "Arrf(" & Fix(X) & "," & Fix(Y) & ").BColor=" & Arrf(Fix(X), Fix(Y), OFP.CurrentPage).BColor
    Debug.Print QBCToAnsiBC(Arrf(Fix(X), Fix(Y), OFP.CurrentPage).BColor)
End If
End Sub



Public Sub DoDraw(ByVal X As Integer, ByVal Y As Integer, ByVal tstr As String)
On Error GoTo out
    Call DoDraw_A(Arrf, X, Y, OFP.CurrentPage, tstr, SysInfo.ForColor)
    Call ShowIt(X, Y)
Exit Sub
out:
    Debug.Print "DoDraw Error Out"
End Sub

Public Sub DoDrawBC(ByVal X As Integer, ByVal Y As Integer)

On Error GoTo out
        Call DoDrawBC_A(Arrf, X, Y, OFP.CurrentPage, SysInfo.BacColor)
        Call ShowIt(X, Y)
Exit Sub
out:
    Debug.Print "DoDrawBC Error Out"

End Sub
Public Sub DoErease(ByVal X As Single, ByVal Y As Single)
On Error GoTo out
    Dim tmpInt As Integer
    tmpInt = DoErease_A(Arrf, X, Y, OFP.CurrentPage)
    Call ShowIt(X, Y)
    If tmpInt <> 0 Then Call ShowIt(X + tmpInt, Y)
'Pic1.Refresh
Exit Sub
out:
Debug.Print "DoErease Error Out"
End Sub
Public Sub DoEreaseB(ByVal X As Integer, ByVal Y As Integer)
    On Error GoTo out

    Dim tmpInt As Integer
    tmpInt = DoEreaseB_A(Arrf, X, Y, OFP.CurrentPage)
    Call ShowIt(X, Y)
    If tmpInt <> 0 Then Call ShowIt(X + tmpInt, Y)
    Call ShowIt(X, Y)
Exit Sub
out:
Debug.Print "DoEreaseB Error Out"
End Sub

Public Sub DoEreaseAll()
    Dim tmpCL As ColorLayer
    tmpCL.Ansi = 0
    tmpCL.BColor = 0
    tmpCL.Color = 7
    For i = 0 To UBound(Arrf, 2)
        For j = 0 To UBound(Arrf, 1)
            Arrf(j, i, OFP.CurrentPage) = tmpCL
            DoEvents
        Next j
    Next i
    Pic1.Cls
    AD.SetTraget
End Sub

Public Sub ReDraw()
    Pic1.Cls
    AD.SetTraget
    For i = 0 To UBound(Arrf, 2)
        For j = 0 To UBound(Arrf, 1)
            Call ShowIt(j, i)
            DoEvents
        Next j
    Next i
    Pic1.Refresh
End Sub
Public Sub ReDraw_Area(ByVal x1 As Integer, ByVal Y1 As Integer, ByVal X2 As Integer, ByVal Y2 As Integer)
    For i = Y1 To Y2
        For j = x1 To X2
            Call ShowIt(j, i)
            DoEvents
        Next j
        DoEvents
    Next i
    Pic1.Refresh
End Sub
Public Sub SetColorBoard()
    '設置調色盤
    For i = 0 To 15
'        Pic2.Line (i Mod 2, Fix(i / 2))-((i Mod 2) + 1, Fix(i / 2) + 1), QBColor(i), BF
    Next i

    For i = 0 To 7
'        Pic3.Line (i Mod 2, Fix(i / 2))-((i Mod 2) + 1, Fix(i / 2) + 1), QBColor(i), BF
    Next i
End Sub
Public Sub GetForColor(ByVal X As Integer, ByVal Y As Integer)

    SysInfo.ForColor = X + 2 * Y
    'Debug.Print "SysInfo.ForColor=" & SysInfo.ForColor
    Shape2.BackColor = QBColor(SysInfo.ForColor)
    Text1.ForeColor = QBColor(SysInfo.ForColor)
    
End Sub
Public Sub GetBackColor(ByVal X As Integer, ByVal Y As Integer)

    SysInfo.BacColor = X + 2 * Y
    'Debug.Print "SysInfo.BacColor=" & SysInfo.BacColor
    Shape1.BackColor = QBColor(SysInfo.BacColor)
    Text1.BackColor = QBColor(SysInfo.BacColor)

End Sub

Public Sub SetSize(W As Integer, H As Integer, Z As Integer, filetype As Byte)
    Pic1.ScaleMode = 1
    Pic1.Height = H * SysInfo.PPA + 30
    Pic1.Width = (W / 2) * SysInfo.PPA + 30
    'Debug.Print "H * SysInfo.PPA + 30=" & (H * SysInfo.PPA + 30)
    Frame5.Width = Pic1.Width + 240
    Frame5.Height = Pic1.Height + 500
    
'    If Frame5.Top + Frame5.Height > ToolP.Top + ToolP.Height Then
'        Frame4.Top = Frame5.Height + Frame5.Top + 100
        'Frame6.Top = Frame5.Height + Frame5.Top + 1500
        'Debug.Print "Frame4.Top=" & Frame4.Top
    '
'        Frame4.Top = ToolP.Top + ToolP.Height + 100
        'Frame6.Top = 5640
    '
    '設定工具屬性的位置
    Call SetToolPPos
    
    If Me.WindowState = 0 Then
'        me.Height = Frame4.Top + Frame4.Height + 850
'        me.Width = ToolP.Left + ToolP.Width + 150
        Me.Left = (Screen.Width - Me.Width) \ 2
        Me.Top = (Screen.Height - Me.Height) \ 2
        'If me.Width < Frame5.Left + Frame5.Width + 150 Then me.Width = Frame5.Left + Frame5.Width + 150
    End If
    'Debug.Print "H=" & H
    Pic1.ScaleHeight = H
    Pic1.ScaleWidth = W
    
    'Debug.Print "Pic1.ScaleHeight=" & Pic1.ScaleHeight
    'Debug.Print "Pic1.ScaleWidth=" & Pic1.ScaleWidth
    ReDim Arrf(0 To W - 1, 0 To H - 1, 1 To Z) As ColorLayer
    Call ArrfPreValue
    '設定檔案屬性
    OFP.filetype = filetype
    '設定繪圖物件的scale單位
    OFP.CurrentPage = 1
    
    'AD.TwipsPerScaleX = 285 / (2 * Screen.TwipsPerPixelX)
    'AD.TwipsPerScaleY = 285 / Screen.TwipsPerPixelY
End Sub

Public Sub CreatAnsiTxt_Area(ByRef Ansitxt As String, x1 As Integer, Y1 As Integer, X2 As Integer, Y2 As Integer)

    Dim AnsiLine As String
    Dim tmpCCC As New ColorCodeCreater
    Dim maxX As Integer
    Dim maxY As Integer
    Dim maxZ As Integer
    maxX = UBound(Arrf, 1)
    maxY = UBound(Arrf, 2)
    maxZ = UBound(Arrf, 3)
    On Error Resume Next
    Ansitxt = "[m"

        For i = Y1 To Y2
            AnsiLine = ""
            For j = x1 To X2
                If Arrf(j, i, OFP.CurrentPage).Ansi <> -1 Then
                    AnsiLine = AnsiLine & tmpCCC.GetCode(Arrf(j, i, OFP.CurrentPage).Ansi, Arrf(j, i, OFP.CurrentPage).Color, Arrf(j, i, OFP.CurrentPage).BColor)
                End If
                DoEvents
            Next j
            If tmpCCC.preBColor = 0 Then
                Ansitxt = Ansitxt & RTrim(AnsiLine) & vbCrLf
            Else
                Ansitxt = Ansitxt & AnsiLine & vbCrLf
            End If
            'tmpCCC.Clear
        Next i
            Ansitxt = Ansitxt & "[m"
            tmpCCC.Clear

    'Text3.text = Ansitxt
End Sub
Public Sub CreatAnsiTxt_VAA_v3(ByRef Ansitxt As String, Optional ByVal AddLine As String)

    Dim AnsiLine As String
    Dim the23line As String
    Select Case OFP.filetype
        Case 1
            the23line = ""
        Case 2
            the23line = AddLine & vbCrLf
        Case 3
            the23line = ""
    End Select
    Dim tmpCCC As New ColorCodeCreater
    Dim maxX As Integer
    Dim maxY As Integer
    Dim maxZ As Integer
    maxX = UBound(Arrf, 1)
    maxY = UBound(Arrf, 2)
    maxZ = UBound(Arrf, 3)
    On Error Resume Next
    Dim tmpL As Long
    Ansitxt = "[m"
    For k = 1 To maxZ
        For i = 0 To maxY
            AnsiLine = ""
            For j = 0 To maxX
                If Arrf(j, i, k).Ansi <> -1 Then
                    AnsiLine = AnsiLine & tmpCCC.GetCode(Arrf(j, i, k).Ansi, Arrf(j, i, k).Color, Arrf(j, i, k).BColor)
                End If
                DoEvents
            Next j
            If tmpCCC.preBColor = 0 Then
                Ansitxt = Ansitxt & RTrim(AnsiLine) & vbCrLf
            Else
                Ansitxt = Ansitxt & AnsiLine & vbCrLf
            End If
            'tmpCCC.Clear
            tmpL = tmpL + 1
        Next i
            Ansitxt = Ansitxt & "[m" & k & "-" & tmpL & the23line
            tmpCCC.Clear
    Next k
    'Text3.text = Ansitxt
End Sub
Public Sub CreatAnsiTxt_VAA(ByRef Ansitxt As String, Optional ByVal AddLine As String)  '舊版的編譯函式
    'Dim Ansitxt As String
    Dim AnsiLine As String
    Dim the23line As String
    Select Case OFP.filetype
        Case 1
            the23line = ""
        Case 2
            the23line = AddLine & vbCrLf
        Case 3
            the23line = ""
    End Select
On Error Resume Next
    For k = 1 To UBound(Arrf, 3)
        For i = 0 To UBound(Arrf, 2)
            AnsiLine = ""
            Select Case Arrf(0, i, k).Ansi
                Case Is = 0 And 32
                    Arrf(0, i, k).Color = 7
                    If Arrf(0, i, k).BColor = 0 Then
                        AnsiLine = AnsiLine & Chr(32)
                    Else
                        If Arrf(0, i, k).BColor = 0 Then
                            AnsiLine = AnsiLine & "[m" & Chr(32)
                        Else
                            AnsiLine = AnsiLine & "[" & QBCToAnsiBC(Arrf(0, i, k).BColor) & "m" & Chr(32)
                        End If
                    End If
                    
                Case Is = -1
            
                Case Else
                    If (Arrf(0, i, k).Color = 7) And (Arrf(0, i, k).BColor = 0) Then
                
                        AnsiLine = AnsiLine & Chr(Arrf(0, i, k).Ansi)
                    Else
                
                        If QBCToAnsiC(Arrf(0, i, k).Color) = "" Then
                            AnsiLine = AnsiLine & "[" & QBCToAnsiBC(Arrf(0, i, k).BColor) & QBCToAnsiC(Arrf(0, i, k).Color) & "m" & Chr(Arrf(0, i, k).Ansi)
                        Else
                            AnsiLine = AnsiLine & "[" & QBCToAnsiBC(Arrf(0, i, k).BColor) & ";" & QBCToAnsiC(Arrf(0, i, k).Color) & "m" & Chr(Arrf(0, i, k).Ansi)
                        End If
                    End If
            End Select
        For j = 1 To UBound(Arrf, 1)
            Select Case Arrf(j, i, k).Ansi
                Case Is = 0 And 32
                    Arrf(j, i, k).Color = 7
                    If Arrf(j, i, k).BColor = Arrf(j - 1, i, k).BColor Then
                        AnsiLine = AnsiLine & Chr(32)
                    Else
                        If Arrf(j, i, k).BColor = 0 Then
                            AnsiLine = AnsiLine & "[m" & Chr(32)
                        Else
                            AnsiLine = AnsiLine & "[" & QBCToAnsiBC(Arrf(j, i, k).BColor) & "m" & Chr(32)
                        End If
                    End If
                        
                Case Is = -1
                
                Case Else
                    If (Arrf(j, i, k).Color = Arrf(j - 1, i, k).Color) And (Arrf(j, i, k).BColor = Arrf(j - 1, i, k).BColor) And (Arrf(j - 1, i, k).Ansi <> 0) Then
                
                        AnsiLine = AnsiLine & Chr(Arrf(j, i, k).Ansi)
                    Else
                
                        If Arrf(j, i, k).Color = 7 And Arrf(j, i, k).BColor = 0 Then
                            AnsiLine = AnsiLine & "[m" & Chr(Arrf(j, i, k).Ansi)
                        Else
                            AnsiLine = AnsiLine & "[" & QBCToAnsiC(Arrf(j, i, k).Color) & ";" & QBCToAnsiBC(Arrf(j, i, k).BColor) & "m" & Chr(Arrf(j, i, k).Ansi)
                        End If
                    End If
            End Select
        'Debug.Print "Arrf(" & j & "," & i & ").Ansi=" & Arrf(j, i, k).Ansi
            DoEvents
        Next j
        Ansitxt = Ansitxt & "[m" & RTrim(AnsiLine) & vbCrLf
        DoEvents
        Next i
    '第23行的重複
    Ansitxt = Ansitxt & the23line
Next k
'Text3.text = Ansitxt

End Sub
Public Sub CreatAnsiTxt_NoColor(ByRef Ansitxt As String)
'Dim Ansitxt As String
    Dim AnsiLine As String
    Dim the23line As String
    Select Case OFP.filetype
        Case 3
            the23line = ""
        Case 2
            the23line = vbCrLf
        Case 3
            the23line = ""
    End Select
    On Error Resume Next
    For k = 1 To UBound(Arrf, 3)
        For i = 0 To UBound(Arrf, 2)
            AnsiLine = ""
        
            For j = 0 To UBound(Arrf, 1)
                Select Case Arrf(j, i, k).Ansi
                    Case Is = 0
                            AnsiLine = AnsiLine & Chr(32)
                    Case Is = -1
                    
                    Case Else
                            AnsiLine = AnsiLine & Chr(Arrf(j, i, k).Ansi)
                End Select
                'Debug.Print "Arrf(" & j & "," & i & ").Ansi=" & Arrf(j, i, k).Ansi
                DoEvents
            Next j
            Ansitxt = Ansitxt & RTrim(AnsiLine) & vbCrLf
            DoEvents
        Next i
        '第23行的重複
        Ansitxt = Ansitxt & the23line
    Next k
    'Text3.text = Ansitxt
    
End Sub
Public Sub SetNowAnsi()
    On Error GoTo out
    If Option1.value = True Then
        NowAnsi = Text1.text
        Toolbar1.Buttons(4).Enabled = True
        SysInfo.ForeSource = 1
        SysInfo.LastAnsi = Asc(NowAnsi)
    End If
    If Option2.value = True Then
        Toolbar1.Buttons(4).Enabled = False
        If SysInfo.EdMode = 4 Then
            Toolbar1.Buttons(1).value = tbrPressed
            SysInfo.EdMode = 1
        End If
    
        SysInfo.ForeSource = 2
    End If

    If Option3.value = True Then
        NowAnsi = Text6.text
        Toolbar1.Buttons(4).Enabled = True
        
        SysInfo.ForeSource = 3
        SysInfo.SSAnsi = Asc(NowAnsi)
    End If
Exit Sub
out:
    Debug.Print "SetNowAnsi Error Out"
End Sub

Public Sub DoMutiDraw(ByVal X As Integer, ByVal Y As Integer, ByVal InAnsis As String, Optional ByVal Mode As Byte)
    Dim Pointer As Integer
    'Debug.Print "InAnsis=" & InAnsis
    Dim tmpstrA() As String
    tmpstrA = Split(InAnsis, vbCrLf)
    For j = 0 To UBound(tmpstrA)
        Pointer = 0

        For i = 1 To Len(tmpstrA(j))
            tstr = Mid(tmpstrA(j), i, 1)
            If Mode = 1 Then
                Call DoErease(X + Pointer, Y + j)
            End If
             Call DoDraw(X + Pointer, Y + j, tstr)
             Debug.Print j & "-" & i & ":" & Asc(tstr) & ";" & tstr
            If Check2.value = 1 Then Call DoDrawBC(X + Pointer, Y + j)
            Pointer = Pointer + Tlen(tstr)
            DoEvents
        Next i
    Next j
End Sub

Public Sub ShowIt(ByVal X As Single, ByVal Y As Single)
On Error GoTo out
    
    If Arrf(X, Y, OFP.CurrentPage).Ansi = -1 Then
        'AD.PrintText "▌", QBColor(Arrf(X, Y, OFP.CurrentPage).BColor), X - 1, Y
        AD.DrawRectangle X - 1, Y, X - 1, Y, Arrf(X, Y, OFP.CurrentPage).BColor
    End If
    'AD.PrintText "▌", QBColor(Arrf(X, Y, OFP.CurrentPage).BColor), X, Y
    AD.DrawRectangle X, Y, X, Y, Arrf(X, Y, OFP.CurrentPage).BColor
    If X <> UBound(Arrf, 1) Then
        If Arrf(X + 1, Y, OFP.CurrentPage).Ansi = -1 Then
            'AD.PrintText "▌", QBColor(Arrf(X, Y, OFP.CurrentPage).BColor), X + 1, Y
            AD.DrawRectangle X + 1, Y, X + 1, Y, Arrf(X, Y, OFP.CurrentPage).BColor
        End If
    End If
    If Arrf(X, Y, OFP.CurrentPage).Ansi = -1 Then
        If Arrf(X, Y, OFP.CurrentPage).Ansi <> 0 Then AD.PrintText Chr(Arrf(X - 1, Y, OFP.CurrentPage).Ansi), QBColor(Arrf(X - 1, Y, OFP.CurrentPage).Color), X - 1, Y
        'If Arrf(X, Y, OFP.CurrentPage).Ansi <> 0 Then AD.PrintText Chr(Arrf(X - 1, Y, OFP.CurrentPage).Ansi), Arrf(X - 1, Y, OFP.CurrentPage).Color, X - 1, Y
    Else
        If Arrf(X, Y, OFP.CurrentPage).Ansi <> 0 Then AD.PrintText Chr(Arrf(X, Y, OFP.CurrentPage).Ansi), QBColor(Arrf(X, Y, OFP.CurrentPage).Color), X, Y
        'If Arrf(X, Y, OFP.CurrentPage).Ansi <> 0 Then AD.PrintText Chr(Arrf(X, Y, OFP.CurrentPage).Ansi), Arrf(X, Y, OFP.CurrentPage).Color, X, Y
    End If

Exit Sub
out:
    Debug.Print "ShowIt Error Out"
    Debug.Print "Error In:" & "(" & X & "," & Y & ")"
End Sub

Public Sub Setoolbar()
    If Check2.value = 1 And Check1.value = 0 Then
        Toolbar1.Buttons(3).Enabled = False
    Else
        Toolbar1.Buttons(3).Enabled = True
    End If
End Sub

Public Sub PaintColor(ByVal X As Integer, ByVal Y As Integer)
On Error GoTo out
    Call PaintColor_A(Arrf, X, Y, OFP.CurrentPage, SysInfo.ForColor, SysInfo.BacColor, Check1.value, Check2.value)
    Call ShowIt(X, Y)
    Pic1.Refresh
Exit Sub
out:

Debug.Print "Paint Color Error Out"
Resume Next
End Sub


Private Sub SetToolP()
Dim preTP As Frame
Set preTP = ToolP

Select Case SysInfo.EdMode
Case Is = 1
    Set ToolP = Frame3
Case Is = 2
    Set ToolP = Frame8
Case Is = 3
'    Set ToolP = Frame3
Case Is = 4
    Set ToolP = Frame3
Case Is = 5
    Set ToolP = Frame10
Case Is = 6
    Set ToolP = Frame9
Case Is = 8
    Set ToolP = Frame7
Case Is = 9
    Set ToolP = Frame11
End Select
'preTP.Visible = False
'ToolP.Left = preTP.Left
'ToolP.Top = preTP.Top
'ToolP.Visible = True
End Sub

Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index

    Case Is = 1
        SysInfo.EdMode = 6
    Case Is = 2
        SysInfo.EdMode = 7
        If List1.ListCount <> 0 And List1.ListIndex = -1 Then List1.ListIndex = 0
    Case Is = 3
    

End Select
If SysInfo.EdMode = 7 Then
    Command6(0).Enabled = False
'    Command6(1).Enabled = False
'    Command6(2).Enabled = False

Else
    Command6(0).Enabled = True
End If

End Sub

Public Sub Set_FileType_Visual()
If OFP.filetype = 1 Then
    Command7(0).Visible = False
    Command7(1).Visible = False
    Command12.Visible = False
    Command9.Visible = False
    Command10.Visible = False
    Command11.Visible = False
    Label10.Visible = False
    Label11.Visible = False
    Text7.Visible = False
    Text8.Visible = False
    Label9.Visible = False
    Combo1.Visible = False
Else
    Command7(0).Visible = True
    Command7(1).Visible = True
    Command12.Visible = True
    Command9.Visible = True
    Command10.Visible = True
    Command11.Visible = True
    Label10.Visible = True
    Label11.Visible = True
    Text7.Visible = True
    Text8.Visible = True
    Label9.Visible = True
    Combo1.Visible = True
End If
If OFP.filetype = 2 Then    '當檔案類型為動畫時開啟的功能
    Command14.Visible = True
    Me_Display.Enabled = True
    Me_Director.Enabled = True
Else
    Command14.Visible = False
    Me_Display.Enabled = False
    Me_Director.Enabled = False
End If

End Sub

Public Sub Set_VAA_Combo()
Combo1.Clear
For i = 1 To UBound(Arrf, 3)
    Combo1.AddItem "第" & i & "頁"


Next i

End Sub



Public Sub SetMoninter()

End Sub

Public Sub ForScreenSize(ByVal Frontsize As Byte)

Select Case Frontsize
    Case 12
        SysInfo.PPA = 225
    Case 13
        SysInfo.PPA = 255
    Case 14
        SysInfo.PPA = 285
    Case 15
        SysInfo.PPA = 315

End Select
Pic1.FontSize = Frontsize

'Pic1.FontSize = 13
'SysInfo.PPA = 255
'Pic1.FontSize = 12
'SysInfo.PPA = 225


'On Error Resume Next
'For Each object In me
'Debug.Print object.Name & ".FontSize=" & object.FontSize
'object.FontSize = 8
'Next

End Sub

Public Sub DisplayVAA()
If Timer1.Interval = 0 Then
    Timer1.Interval = 1000
    Command14.Caption = "停止撥放"
    Me_Display.Caption = "停止撥放"
Else
    Command14.Caption = "撥放"
    Me_Display.Caption = "撥放"
    Timer1.Interval = 0
End If
End Sub

Public Sub SelectAnsi(ByVal X As Integer, ByVal Y As Integer)
Dim dX As Integer
dX = 0
If Arrf(X, Y, OFP.CurrentPage).Ansi = -1 Then
    X = X - 1
    dX = 1
End If
If X < UBound(Arrf, 1) Then
    If Arrf(X + 1, Y, OFP.CurrentPage).Ansi = -1 Then
        dX = 1
    End If
End If
SL.StartPoint_X = X
SL.StartPoint_Y = Y
SL.EndPoint_X = X + dX
SL.EndPoint_Y = Y
SL.DrawSelect
End Sub

Public Sub SetConfic()
Call WriteConfic(App.Path & "\confic.cfg", SysInfo)
End Sub
Public Sub GetConfic()
    Dim ConficDS As SysEnv
    ConficDS = ReadConfic(App.Path & "\confic.cfg")
    '讀取上次離開時所選擇的工具
    Select Case ConficDS.EdMode
        Case 0
        Case 1 To 5
            SysInfo.EdMode = ConficDS.EdMode
'            Toolbar1.Buttons(SysInfo.EdMode).value = tbrPressed
'            Toolbar2.Buttons(1).value = tbrPressed
        Case 6
            
        Case 7
    
    End Select
    
    '上次離開時所使用之筆觸
    On Error GoTo engout
    If ConficDS.SSAnsi <> 0 Then
        Text6.text = Chr(ConficDS.SSAnsi)
    End If
    'Debug.Print "ConficDS.SSAnsi" & ConficDS.SSAnsi
    If ConficDS.LastAnsi <> 0 Then
        
        Text1.text = Chr(ConficDS.LastAnsi)
        'Debug.Print "ConficDS.LastAnsi=" & ConficDS.LastAnsi
    End If
    '上次離開時所選擇的前景來源
    Select Case ConficDS.ForeSource
        Case 0
        Case 1
            Option1.value = True
        Case 2
            Option2.value = True
        Case 3
            Option3.value = True
    End Select
    
    
    If ConficDS.Frontsize <> 0 Then SysInfo.Frontsize = ConficDS.Frontsize
    If ConficDS.ForColor <> 0 Then
        SysInfo.ForColor = ConficDS.ForColor
        'Debug.Print "SysInfo.ForColor=" & SysInfo.ForColor
        Shape2.BackColor = QBColor(SysInfo.ForColor)
        Text1.ForeColor = QBColor(SysInfo.ForColor)
    End If
    If ConficDS.BacColor <> 0 Then
        SysInfo.BacColor = ConficDS.BacColor
        'Debug.Print "SysInfo.BacColor=" & SysInfo.BacColor
        Shape1.BackColor = QBColor(SysInfo.BacColor)
        Text1.BackColor = QBColor(SysInfo.BacColor)
    End If
    '關閉一個檔案時確認儲存
    SysInfo.CheckSave = ConficDS.CheckSave
    '隱藏選擇框
    SysInfo.HideSelect = ConficDS.HideSelect
    '工具屬性置底
    SysInfo.ToolPBoxDown = ConficDS.ToolPBoxDown
Exit Sub
engout:
    MsgBox "請將您的windows地區選項改為中文(台灣)本程式才能正常運作 (特別感謝superlubu呂布提供除錯)"
    Resume Next
End Sub

Public Sub AskSave()
If OFP.IsChanged = True Then
        If MsgBox("是否要在關閉前儲存現在這個檔案" & vbCrLf & OFP.filepath, 36, "提醒") = 6 Then
            Call Me_Save_Click
        End If
End If
End Sub
'>>>>>>>>>>設定視窗標題
Public Sub SetFormCaption(Optional str As String)
Dim tmpStr As String
tmpStr = "Visual Ansi " & App.Major & "." & App.Minor & "." & App.Revision & "a"
If str <> "" Then

    Me.Caption = tmpStr & " - " & str
Else

    If OFP.filepath = "" Then
        Me.Caption = tmpStr
    Else
        If OFP.IsChanged = True Then
            Me.Caption = tmpStr & " - " & FileSys.GetFileName(OFP.filepath) & " * "
        Else
            Me.Caption = tmpStr & " - " & FileSys.GetFileName(OFP.filepath)
        End If

    End If
End If
End Sub

Public Sub ArrfPreValue()
For k = 1 To UBound(Arrf, 3)
    For j = 0 To UBound(Arrf, 2)
        For i = 0 To UBound(Arrf, 1)
            Arrf(i, j, k).Color = 7
            DoEvents
        Next i
    Next j
Next k
End Sub
'>>>>>>>>>>顏色置換功能
Public Sub ExChColor_SetFColor(ByVal X As Integer, ByVal Y As Integer)
If SysInfo.ExChColor.CurrentSel < 2 Then
    SysInfo.ExChColor.Color(SysInfo.ExChColor.CurrentSel) = X + 2 * Y
    Lb_ExCh_Color(SysInfo.ExChColor.CurrentSel).BackColor = QBColor(SysInfo.ExChColor.Color(SysInfo.ExChColor.CurrentSel))
End If

End Sub
Public Sub ExChColor_SetBColor(ByVal X As Integer, ByVal Y As Integer)
If SysInfo.ExChColor.CurrentSel > 1 Then
    SysInfo.ExChColor.Color(SysInfo.ExChColor.CurrentSel) = X + 2 * Y
     Lb_ExCh_Color(SysInfo.ExChColor.CurrentSel).BackColor = QBColor(SysInfo.ExChColor.Color(SysInfo.ExChColor.CurrentSel))
    
End If
End Sub
Public Sub ExChColor_Draw(ByVal X As Integer, ByVal Y As Integer)
On Error GoTo out
'Debug.Print "SysInfo.ExChColor.Color(0)=" & SysInfo.ExChColor.Color(0) & vbCrLf & "Arrf(X, Y, OFP.CurrentPage).Color=" & Arrf(X, Y, OFP.CurrentPage).Color
    Call ExChColor_Draw_A(Arrf, X, Y, OFP.CurrentPage, SysInfo.ExChColor.Color(0), SysInfo.ExChColor.Color(2), SysInfo.ExChColor.Color(1), SysInfo.ExChColor.Color(3), Check1.value, Check2.value)
    Call ShowIt(X, Y)
    Pic1.Refresh
Exit Sub
out:

Debug.Print "ExChColor_Draw  Error Out"
Resume Next
End Sub


Public Sub SetToolPPos()
On Error GoTo out
If SysInfo.ToolPBoxDown = 1 Then
    
    ToolP.Left = Frame4.Left + Frame4.Width
    ToolP.Top = Frame4.Top
    If Me.Height < ToolP.Top + ToolP.Height + 850 Then Me.Height = ToolP.Top + ToolP.Height + 850
Else
    If Frame5.Left + Frame5.Width >= Frame1.Left + Frame1.Width Then
        ToolP.Left = Frame5.Left + Frame5.Width
        'Debug.Print "ToolP.Left=" & ToolP.Left
    Else
        ToolP.Left = Frame1.Left + Frame1.Width

    End If
End If
Exit Sub
out:
End Sub

Public Sub DoErease_Area_In(ByVal tmpX1 As Integer, ByVal tmpY1 As Integer, ByVal tmpX2 As Integer, ByVal tmpY2 As Integer)
    On Error GoTo out
            For j = tmpY1 To tmpY2
                
                For i = tmpX1 To tmpX2
                    If Not (i = tmpX1 And Arrf(tmpX1, j, OFP.CurrentPage).Ansi = -1) And Not (i = tmpX2 And Arrf(tmpX2 + 1, j, OFP.CurrentPage).Ansi = -1) Then
                        Call DoErease_A(Arrf, i, j, OFP.CurrentPage)
                        Call DoEreaseB_A(Arrf, i, j, OFP.CurrentPage)
                        DoEvents
                    End If

                Next i
            Next j
            Call ReDraw_Area(tmpX1 - 1, tmpY1, tmpX2 + 1, tmpY2)
Exit Sub
out:
End Sub
Public Sub DoErease_Area(ByVal tmpX1 As Integer, ByVal tmpY1 As Integer, ByVal tmpX2 As Integer, ByVal tmpY2 As Integer)
    On Error GoTo out
            For j = tmpY1 To tmpY2
                
                For i = tmpX1 To tmpX2
                    
                        Call DoErease_A(Arrf, i, j, OFP.CurrentPage)
                        Call DoEreaseB_A(Arrf, i, j, OFP.CurrentPage)
                        DoEvents
 

                Next i
            Next j
            Call ReDraw_Area(tmpX1 - 1, tmpY1, tmpX2 + 1, tmpY2)
Exit Sub
out:
End Sub

