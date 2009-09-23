VERSION 5.00
Begin VB.Form Form13 
   BorderStyle     =   4  '單線固定工具視窗
   Caption         =   "前景來源 - 自訂文句"
   ClientHeight    =   2805
   ClientLeft      =   45
   ClientTop       =   270
   ClientWidth     =   4065
   LinkTopic       =   "Form13"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2805
   ScaleWidth      =   4065
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所屬視窗中央
   Begin VB.CommandButton Command3 
      Caption         =   "關閉"
      Height          =   270
      Left            =   2985
      TabIndex        =   4
      Top             =   2415
      Width           =   720
   End
   Begin VB.CommandButton Command2 
      Caption         =   "清除"
      Height          =   270
      Left            =   2025
      TabIndex        =   3
      Top             =   2400
      Width           =   720
   End
   Begin VB.CommandButton Command1 
      Caption         =   "套用"
      Height          =   285
      Index           =   1
      Left            =   1110
      TabIndex        =   2
      Top             =   2400
      Width           =   780
   End
   Begin VB.CommandButton Command1 
      Caption         =   "編輯"
      Height          =   285
      Index           =   0
      Left            =   210
      TabIndex        =   1
      Top             =   2400
      Width           =   780
   End
   Begin VB.TextBox Text1 
      Height          =   2280
      Left            =   30
      MultiLine       =   -1  'True
      ScrollBars      =   3  '兩者皆有
      TabIndex        =   0
      Top             =   30
      Width           =   4005
   End
End
Attribute VB_Name = "Form13"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click(Index As Integer)
Dim tmpbyte As Byte
tmpbyte = 1 - Index
Text1.Enabled = tmpbyte
Command1(tmpbyte).Enabled = True
Command1(Index).Enabled = False
Select Case Index
    Case 0
        Command2.Enabled = True
    Case 1
        Command2.Enabled = False
        FString.str = Text1.text
        Dim tmpstrA() As String
        Dim tmpCounter As Integer
        tmpstrA = Split(FString.str, vbCrLf)
        FString.StrLen(1) = UBound(tmpstrA) + 1
        For i = 0 To UBound(tmpstrA)
            
            Replace tmpstrA(i), vbCrLf, "" 'fix for 2000
            If tmpCounter < Tlen(tmpstrA(i)) Then
                tmpCounter = Tlen(tmpstrA(i))
            End If
        Next i
        FString.StrLen(0) = tmpCounter
End Select
End Sub

Private Sub Command2_Click()

Text1.text = ""
End Sub

Private Sub Command3_Click()
    On Error GoTo out
    Unload Me
    Exit Sub
out:
End Sub

Private Sub Form_Load()
Text1.text = FString.str
Call Command1_Click(0)
End Sub


