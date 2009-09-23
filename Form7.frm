VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form7 
   Caption         =   "匯入文章/彩色碼"
   ClientHeight    =   5325
   ClientLeft      =   240
   ClientTop       =   345
   ClientWidth     =   8745
   Icon            =   "Form7.frx":0000
   LinkTopic       =   "Form7"
   ScaleHeight     =   5325
   ScaleWidth      =   8745
   StartUpPosition =   3  '系統預設值
   Begin VB.Frame Frame3 
      Caption         =   "提醒"
      Height          =   810
      Left            =   3630
      TabIndex        =   16
      Top             =   4485
      Width           =   4830
      Begin VB.Label Label2 
         Caption         =   "除了使用本功能之外，可以使用複製貼上的""從Windows剪貼簿擷取""功能來貼上。"
         Height          =   495
         Left            =   315
         TabIndex        =   17
         Top             =   240
         Width           =   4185
      End
   End
   Begin VB.Frame Frame2 
      Height          =   645
      Index           =   1
      Left            =   60
      TabIndex        =   11
      Top             =   3975
      Width           =   3405
      Begin VB.CheckBox Check2 
         Caption         =   "去背"
         Height          =   255
         Left            =   1965
         TabIndex        =   15
         Top             =   270
         Width           =   1260
      End
      Begin VB.CommandButton Command4 
         Caption         =   "彩色貼上"
         Height          =   315
         Left            =   210
         TabIndex        =   14
         Top             =   225
         Width           =   1440
      End
   End
   Begin VB.Frame Frame2 
      Height          =   630
      Index           =   0
      Left            =   75
      TabIndex        =   10
      Top             =   4665
      Width           =   3420
      Begin VB.CheckBox Check1 
         Caption         =   "蓋過原本的"
         Height          =   195
         Left            =   1965
         TabIndex        =   13
         Top             =   270
         Width           =   1245
      End
      Begin VB.CommandButton Command1 
         Caption         =   "一般貼上"
         Height          =   315
         Left            =   180
         TabIndex        =   12
         Top             =   210
         Width           =   1440
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "關 閉"
      Height          =   345
      Left            =   7035
      TabIndex        =   8
      Top             =   4080
      Width           =   1140
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  '靠右對齊
      Height          =   300
      Index           =   1
      Left            =   5490
      TabIndex        =   6
      Text            =   "0"
      Top             =   4110
      Width           =   420
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  '靠右對齊
      Height          =   300
      Index           =   0
      Left            =   4815
      TabIndex        =   5
      Text            =   "0"
      Top             =   4110
      Width           =   435
   End
   Begin VB.Frame Frame1 
      Caption         =   "錯誤檢查"
      Height          =   2565
      Left            =   2070
      TabIndex        =   0
      Top             =   735
      Visible         =   0   'False
      Width           =   4890
      Begin VB.CommandButton Command3 
         Caption         =   "硬要貼上"
         Height          =   315
         Index           =   2
         Left            =   2895
         Style           =   1  '圖片外觀
         TabIndex        =   4
         Top             =   1995
         Width           =   1005
      End
      Begin VB.CommandButton Command3 
         Caption         =   "自動斷行"
         Height          =   315
         Index           =   1
         Left            =   375
         Style           =   1  '圖片外觀
         TabIndex        =   3
         Top             =   1980
         Width           =   1005
      End
      Begin VB.CommandButton Command3 
         Caption         =   "我自己改"
         Height          =   315
         Index           =   0
         Left            =   1590
         Style           =   1  '圖片外觀
         TabIndex        =   2
         Top             =   1980
         Width           =   1005
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  '平面
         Height          =   1335
         Left            =   345
         MultiLine       =   -1  'True
         ScrollBars      =   2  '垂直捲軸
         TabIndex        =   1
         Top             =   465
         Width           =   4290
      End
   End
   Begin RichTextLib.RichTextBox Text1 
      Height          =   3930
      Left            =   15
      TabIndex        =   9
      Top             =   0
      Width           =   8670
      _ExtentX        =   15293
      _ExtentY        =   6932
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"Form7.frx":1042
   End
   Begin VB.Label Label1 
      Caption         =   "左上角座標 X              Y"
      Height          =   240
      Left            =   3660
      TabIndex        =   7
      Top             =   4170
      Width           =   1920
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ErrorC() As Integer
Private Sub Command1_Click()
    If InStr(Text1.text, "[") <> 0 Then
        If MsgBox("文字的內容可能包含彩色碼的訊息，建議您改以彩色貼上，否則將無法正確顯示出其原貌 " & vbCrLf & "都這樣說了 你還硬是要貼上嗎?", 65, "偵測到彩色碼") = vbCancel Then Exit Sub
    End If
    Call CheckOver
    OFP.IsChanged = True
    Call Form1.SetFormCaption

End Sub



Private Sub Command2_Click()
Unload Form7
End Sub

Private Sub Command3_Click(Index As Integer)
        Frame1.Visible = False
        Text1.Enabled = True
Select Case Index
    Case 0
    Case 1

        Call InsertCrLf
    Case 2
        Dim a() As String
        Dim x_shift As Integer
        Dim y_shift As Integer
        x_shift = Val(Text3(0).text)
        y_shift = Val(Text3(1).text)
        a = Split(Text1.text, vbCrLf)
        
        For i = 0 To UBound(a)
            Call Form1.DoMutiDraw(x_shift, i + y_shift, a(i), Check1.Value)
            
        Next i
        Form1.Pic1.Refresh
End Select
End Sub

Private Sub Command4_Click()
'On Error GoTo out
    'Dim CCR As New ColorCodeReader
    X = Val(Text3(0).text)
    Y = Val(Text3(1).text)
    Form1.CCR.SetCCStr = Text1.text
    '解析傳進來的彩色碼
    Call Form1.CCR.AnalyzeCC
    '貼上  是否貼上Check2.Value
    Call Form1.CCR.Post2Arrf(X, Y, Check2.Value)
    Call Form1.AD.ReDraw
    '標的內容變更
    OFP.IsChanged = True
    Call Form1.SetFormCaption
Exit Sub
out:
    Debug.Print "匯入彩色碼Error : " & Err.Description
End Sub

Private Sub Form_Load()
'Text1.text = String(80, "L")
Text1.RightMargin = Screen.Width * 2
End Sub


Public Function Tlen(str1 As String)

Tlen = StrConv(str1, vbFromUnicode)

End Function

Public Sub CheckOver()
Dim a() As String
Dim topnum As Integer
Dim ErrorStr As String
Text2.text = ""
topnum = UBound(Arrf, 1) - Val(Text3(0).text) + 1

a = Split(Text1.text, vbCrLf)
ReDim ErrorC(0 To UBound(a))
For i = 0 To UBound(a)
    Debug.Print "a(" & i & ")=" & a(i)
    If LenB(StrConv(a(i), vbFromUnicode)) > topnum Then
        ErrorStr = ErrorStr & " " & (i + 1)
        
    End If
Next i
If Trim(ErrorStr) <> "" Then
    Frame1.Visible = True
    Text1.Enabled = False
    Text2.text = "每行只能容納" & topnum & "個字元(中文字算兩個)" & vbCrLf & "超過的有第" & ErrorStr & "行"
Else
    Dim x_shift As Integer
    Dim y_shift As Integer
    x_shift = Val(Text3(0).text)
    y_shift = Val(Text3(1).text)
    For i = 0 To UBound(a)
        Call Form1.DoMutiDraw(x_shift, i + y_shift, a(i), Check1.Value)
    Next i
    Form1.Pic1.Refresh
End If
End Sub


Private Sub InsertCrLf()
Dim a() As String
Dim tempstr As String
Dim ErrorStr As String
Dim topnum As Integer
topnum = UBound(Arrf, 1) - Val(Text3(0).text) + 1
a = Split(Text1.text, vbCrLf)

For i = 0 To UBound(a)
    Debug.Print "a(" & i & ")=" & a(i)
    If LenB(StrConv(a(i), vbFromUnicode)) > topnum Then
        Call StringCrlf(a(i), topnum)
        
    End If
    tempstr = tempstr & a(i) & vbCrLf
Next i
Text1.text = tempstr
End Sub

Private Sub StringCrlf(ByRef str1 As String, ByVal fixlen As Byte)
Debug.Print "StringCrlf was called"

Dim strarray As String
Dim tempstr As String
strarray = str1
Debug.Print "strarray=" & strarray
str1 = ""
For i = 0 To Len(strarray)
    tempstr = tempstr & Mid(strarray, i + 1, 1)

    If (LenB(StrConv(tempstr, vbFromUnicode)) >= (fixlen - 1) And LenB(StrConv(Mid(strarray, i + 2, 1), vbFromUnicode)) = 2) Or LenB(StrConv(tempstr, vbFromUnicode)) >= fixlen Then
        str1 = str1 & tempstr & vbCrLf
            Debug.Print "lenb of tempstr=" & LenB(StrConv(tempstr, vbFromUnicode))
        tempstr = ""
    End If

    DoEvents

Next i
str1 = str1 & tempstr


End Sub


