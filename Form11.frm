VERSION 5.00
Begin VB.Form Form11 
   BorderStyle     =   4  '��u�T�w�u�����
   Caption         =   "�ʵe�ɼ�"
   ClientHeight    =   4275
   ClientLeft      =   45
   ClientTop       =   270
   ClientWidth     =   5100
   LinkTopic       =   "Form11"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4275
   ScaleWidth      =   5100
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '���ݵ�������
   Begin VB.CommandButton Command8 
      Caption         =   "����"
      Height          =   285
      Left            =   3720
      TabIndex        =   22
      Top             =   3855
      Width           =   1050
   End
   Begin VB.Frame Frame3 
      Caption         =   "�ɶ��b"
      Height          =   885
      Left            =   1170
      TabIndex        =   13
      Top             =   2250
      Width           =   3780
      Begin VB.CommandButton Command7 
         Caption         =   "�]�w"
         Height          =   255
         Left            =   1905
         TabIndex        =   16
         Top             =   225
         Width           =   750
      End
      Begin VB.TextBox Text2 
         Height          =   270
         Left            =   1035
         TabIndex        =   14
         Top             =   210
         Width           =   750
      End
      Begin VB.Label Label6 
         Height          =   225
         Left            =   2460
         TabIndex        =   21
         Top             =   525
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "�`��"
         Height          =   225
         Left            =   2025
         TabIndex        =   20
         Top             =   555
         Width           =   510
      End
      Begin VB.Label Label4 
         Height          =   225
         Left            =   1155
         TabIndex        =   19
         Top             =   555
         Width           =   825
      End
      Begin VB.Label Label3 
         Caption         =   "(�i�h��)"
         Height          =   195
         Left            =   2745
         TabIndex        =   18
         Top             =   255
         Width           =   780
      End
      Begin VB.Label Label2 
         Caption         =   "�w�g�L�ɶ�"
         Height          =   210
         Index           =   1
         Left            =   165
         TabIndex        =   17
         Top             =   570
         Width           =   945
      End
      Begin VB.Label Label2 
         Caption         =   "���d�ɶ�"
         Height          =   210
         Index           =   0
         Left            =   165
         TabIndex        =   15
         Top             =   255
         Width           =   795
      End
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  '����
      Height          =   960
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   3  '��̬Ҧ�
      TabIndex        =   11
      Top             =   3180
      Width           =   3465
   End
   Begin VB.Frame Frame2 
      Caption         =   "�h��ʧ@"
      Height          =   2160
      Left            =   3135
      TabIndex        =   8
      Top             =   75
      Width           =   1815
      Begin VB.CommandButton Command9 
         Caption         =   "�R��"
         Height          =   300
         Left            =   345
         TabIndex        =   23
         Top             =   1095
         Width           =   1140
      End
      Begin VB.CommandButton Command6 
         Caption         =   "����"
         Height          =   300
         Left            =   345
         TabIndex        =   10
         Top             =   675
         Width           =   1155
      End
      Begin VB.CommandButton Command5 
         Caption         =   "�洫"
         Height          =   270
         Left            =   345
         TabIndex        =   9
         Top             =   285
         Width           =   1170
      End
      Begin VB.Label Label1 
         Caption         =   "���� : ����Shift ��Ctrl�t�X�ƹ��i�h�����"
         Height          =   570
         Left            =   135
         TabIndex        =   12
         Top             =   1530
         Width           =   1560
      End
   End
   Begin VB.CheckBox Check1 
      Caption         =   "�w��"
      Height          =   195
      Left            =   3810
      TabIndex        =   7
      Top             =   3180
      Value           =   1  '�֨�
      Width           =   705
   End
   Begin VB.CommandButton Command4 
      Caption         =   "���s��z"
      Height          =   285
      Left            =   3705
      TabIndex        =   6
      Top             =   3480
      Width           =   1050
   End
   Begin VB.Frame Frame1 
      Caption         =   "���ʧ@"
      Height          =   2175
      Left            =   1200
      TabIndex        =   1
      Top             =   60
      Width           =   1860
      Begin VB.CommandButton Command3 
         Caption         =   "���J�ťխ�"
         Height          =   285
         Left            =   240
         TabIndex        =   5
         Top             =   1560
         Width           =   1395
      End
      Begin VB.CommandButton Command2 
         Caption         =   "�U��"
         Height          =   285
         Index           =   1
         Left            =   240
         TabIndex        =   4
         Top             =   1125
         Width           =   1380
      End
      Begin VB.CommandButton Command2 
         Caption         =   "�W��"
         Height          =   285
         Index           =   0
         Left            =   240
         TabIndex        =   3
         Top             =   735
         Width           =   1380
      End
      Begin VB.CommandButton Command1 
         Caption         =   "�R��"
         Height          =   300
         Left            =   225
         TabIndex        =   2
         Top             =   300
         Width           =   1380
      End
   End
   Begin VB.ListBox List1 
      Height          =   2940
      ItemData        =   "Form11.frx":0000
      Left            =   60
      List            =   "Form11.frx":0002
      MultiSelect     =   2  '�i���h�����
      TabIndex        =   0
      Top             =   105
      Width           =   1035
   End
End
Attribute VB_Name = "Form11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SelArr() As Integer
Private Sub Command1_Click()
    Call SetCaption("�R����...")
    Dim preIndex As Integer
    preIndex = List1.ListIndex
    If preIndex = List1.ListCount - 1 Then
        preIndex = preIndex - 1
        If preIndex < 0 Then
            AddMsgStr "�u�Ѥ@���F�A�L�k�R��"
            Call SetCaption
        End If
    End If
        
    DelPage List1.ListIndex + 1
    OFP.IsChanged = True
    Call Form1.SetFormCaption
    Call ListPages
    List1.ListIndex = preIndex
    AddMsgStr "�w�R�� - " & preIndex + 1
    Call SetCaption
    Call Form1.Set_VAA_Combo
End Sub

Private Sub Command2_Click(Index As Integer)
Call SetCaption("���ʤ�...")
Dim tmpInt As Integer
Dim tmpStr As String
tmpInt = List1.ListIndex
Select Case Index
    Case 0
        If tmpInt <> 0 Then

            ExChPage tmpInt + 1, tmpInt
            tmpStr = List1.List(tmpInt)
            List1.List(tmpInt) = List1.List(tmpInt - 1)
            List1.List(tmpInt - 1) = tmpStr
            List1.ListIndex = tmpInt - 1
        End If
    Case 1
        If tmpInt <> List1.ListCount - 1 Then
            ExChPage tmpInt + 1, tmpInt + 2
            tmpStr = List1.List(tmpInt + 1)
            List1.List(tmpInt + 1) = List1.List(tmpInt)
            List1.List(tmpInt) = tmpStr
            List1.ListIndex = tmpInt + 1
        End If
End Select
OFP.IsChanged = True
Call Form1.SetFormCaption
Call SetCaption
End Sub

Private Sub Command3_Click()
Call SetCaption("���J��...")
Dim preIndex As Integer
preIndex = List1.ListIndex
InsertBlank List1.ListIndex + 1
OFP.IsChanged = True
Call Form1.SetFormCaption
Call ListPages
List1.ListIndex = preIndex
Call SetCaption
Call Form1.Set_VAA_Combo
End Sub

Private Sub Command4_Click()
    Call ListPages
End Sub

Private Sub Command5_Click()
    '�洫
    Call SetCaption("�洫��...")
    If List1.selCount <> 2 Then
        Dim tmpStr As String
        Call GetSelect
        ExChPage SelArr(0), SelArr(1)
        '�洫list�W����r SelArr�O�spage ��1��~�i��list��index
        tmpStr = List1.List(SelArr(0) - 1)
        List1.List(SelArr(0) - 1) = List1.List(SelArr(1) - 1)
        List1.List(SelArr(1) - 1) = tmpStr
        AddMsgStr "�w�g�N " & List1.List(SelArr(0) - 1) & "�P" & List1.List(SelArr(1) - 1) & "�����e�洫"
        OFP.IsChanged = True
        Call Form1.SetFormCaption
    Else
        AddMsgStr "�L�k�ާ@ : ���ʧ@�u��������Ӫ��� �ӧA��F" & List1.selCount & "��"
    End If
    Call SetCaption
End Sub

Private Sub Command6_Click()
    Call SetCaption("���त...")
    Dim tmpInt As Integer
    Dim tmpSelCount As Integer
    Dim tmpint2 As Integer
    If List1.selCount <= 1 Then
        AddMsgStr "�L�k�ާ@ : ���������ӥH�W������"
    Else
        tmpSelCount = GetSelect
        tmpInt = tmpSelCount \ 2 - 1
        For i = 0 To tmpInt
            tmpint2 = tmpSelCount - i - 1
            ExChPage SelArr(i), SelArr(tmpint2)
            tmpStr = List1.List(SelArr(i) - 1)
            List1.List(SelArr(i) - 1) = List1.List(SelArr(tmpint2) - 1)
            List1.List(SelArr(tmpint2) - 1) = tmpStr
    
        Next i
        AddMsgStr "�w�N������������ǭ���"
        OFP.IsChanged = True
        Call Form1.SetFormCaption
    End If
    Call SetCaption
End Sub

Private Sub Command7_Click()
    Dim i As Integer
    For i = 0 To List1.ListCount - 1
        If List1.Selected(i) Then
            timeLine(i + 1) = Val(Text2.text)
            
        End If
    
    Next i
    Call timeSum
    
End Sub

Private Sub Command8_Click()
    Unload Me
End Sub

Private Sub Command9_Click()
    '�h���R��
    Dim selCount As Integer
    selCount = GetSelect()
    OFP.IsChanged = True
    Call Form1.SetFormCaption
    For i = selCount - 1 To 0 Step -1
        Call SetCaption("�R����..." & selCount - i & "/" & selCount)
        DelPage SelArr(i)

        Call ListPages
        
        AddMsgStr "�w�R�� - " & SelArr(i)
        
        Call Form1.Set_VAA_Combo
    Next i
    Call SetCaption
End Sub

Private Sub Form_Load()
    Call ListPages
    Call timeSum
End Sub

Private Sub List1_Click()
    OFP.CurrentPage = List1.ListIndex + 1
    If Check1.Value = 1 Then
        Call Form1.AD.ReDraw
        Call Form1.VAA_SetButton
    End If
    Text2.text = timeLine(List1.ListIndex + 1)
    Call timeSumUntileNow(List1.ListIndex + 1)
End Sub

Public Sub ListPages()
    List1.Clear
    For i = 1 To UBound(Arrf, 3)
        List1.AddItem "��" & i & "��"
    Next i
End Sub

Public Function GetSelect() As Integer
    Dim SCounter As Integer
    ReDim SelArr(List1.selCount - 1)
    For i = 0 To List1.ListCount - 1
        If List1.Selected(i) Then
            SelArr(SCounter) = i + 1
            SCounter = SCounter + 1     'SelArr�O�ƾ�+1
        End If
    Next i
    GetSelect = List1.selCount
End Function

Public Sub AddMsgStr(ByVal str As String)
    Text1.text = Text1.text & str & vbCrLf
    Text1.SelStart = Len(Text1.text)
End Sub

Public Sub SetCaption(Optional ByVal str As String)
If str <> "" Then
    Form11.Caption = "�ʵe�ɼ� - " & str
    Frame1.Enabled = False
    Frame2.Enabled = False
Else
    Form11.Caption = "�ʵe�ɼ�"
    Frame1.Enabled = True
    Frame2.Enabled = True
End If

End Sub


Public Sub timeSumUntileNow(ByVal page As Integer)
    Dim tmptime As Double, tmpStr As String, tmpInt As Long
    For i = 1 To page
        tmptime = tmptime + timeLine(i)
    Next i
    If tmptime < 60 Then
        tmpStr = tmptime & "��"
    ElseIf tmptime > 59 And tmptime < 3600 Then
        tmpInt = tmptime \ 60
        tmpStr = tmpInt & "��" & (tmptime - 60 * tmpInt) & "��"
        
    ElseIf tmptime > 3599 Then
        tmpInt = tmptime \ 60
        tmpStr = (tmptime - 60 * tmpInt) & "��"
        tmpStr = tmpInt \ 60 & "�p��" & (tmpInt Mod 60) & "��" & tmpStr
    End If
    Label4.Caption = tmpStr
    
End Sub

Public Sub timeSum()
    Dim tmptime As Double
    For i = 1 To UBound(timeLine)
        tmptime = tmptime + timeLine(i)
    Next i
    Label6.Caption = tmptime
End Sub

