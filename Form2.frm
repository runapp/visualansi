VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   5340
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   5340
   ScaleWidth      =   4680
   StartUpPosition =   3  '¨t²Î¹w³]­È
   Begin VB.CommandButton Command7 
      Caption         =   "select"
      Height          =   525
      Left            =   2820
      TabIndex        =   8
      Top             =   4665
      Width           =   1155
   End
   Begin VB.CommandButton Command6 
      Caption         =   "if"
      Height          =   510
      Left            =   1170
      TabIndex        =   7
      Top             =   4650
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Command5"
      Height          =   3240
      Left            =   165
      TabIndex        =   6
      Top             =   1395
      Width           =   780
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   1365
      Left            =   2685
      TabIndex        =   5
      Top             =   1380
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   1155
      Left            =   990
      TabIndex        =   4
      Top             =   2820
      Width           =   1665
   End
   Begin VB.CommandButton Command4 
      Caption         =   "SET"
      Height          =   645
      Left            =   990
      TabIndex        =   3
      Top             =   2055
      Width           =   1665
   End
   Begin VB.CommandButton Command2 
      Caption         =   "GET"
      Height          =   630
      Left            =   1005
      TabIndex        =   2
      Top             =   1395
      Width           =   1650
   End
   Begin VB.TextBox Text2 
      Height          =   270
      Left            =   90
      TabIndex        =   1
      Text            =   "Key"
      Top             =   960
      Width           =   4515
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   75
      TabIndex        =   0
      Text            =   "HKEY_CURRENT_USER\Software\Valve\CounterStrike\Settings"
      Top             =   435
      Width           =   4545
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function CloseClipboard Lib "user32" () As Long
Private Declare Function OpenClipboard Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function EmptyClipboard Lib "user32" () As Long
Private Declare Function GetClipboardData Lib "user32" (ByVal wFormat As Long) As Long
Private Declare Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
Private Declare Function IsClipboardFormatAvailable Lib "user32" (ByVal wFormat As Long) As Long
Public Enum EPredefinedClipboardFormatConstants
     CF_TEXT = 1
     CF_BITMAP = 2
     CF_METAFILEPICT = 3
     CF_SYLK = 4
     CF_DIF = 5
     CF_TIFF = 6
     CF_OEMTEXT = 7
     CF_DIB = 8
     CF_PALETTE = 9
     CF_PENDATA = 10
     CF_RIFF = 11
     CF_WAVE = 12
     CF_UNICODETEXT = 13
     CF_ENHMETAFILE = 14
''#if(WINVER >= 0x0400)
     CF_HDROP = 15
     CF_LOCALE = 16
     CF_MAX = 17
'#endif /* WINVER >= 0x0400 */
     CF_OWNERDISPLAY = &H80
     CF_DSPTEXT = &H81
     CF_DSPBITMAP = &H82
     CF_DSPMETAFILEPICT = &H83
     CF_DSPENHMETAFILE = &H8E
'/*
' * "Private" formats don't get GlobalFree()'d
' */
     CF_PRIVATEFIRST = &H200
     CF_PRIVATELAST = &H2FF
'/*
' * "GDIOBJ" formats do get DeleteObject()'d
' */
     CF_GDIOBJFIRST = &H300
     CF_GDIOBJLAST = &H3FF
End Enum
'API memory ±`¼Æ
Private Const GMEM_FIXED = &H0
Private Const GMEM_MOVEABLE = &H2
Private Const GMEM_NOCOMPACT = &H10
Private Const GMEM_NODISCARD = &H20
Private Const GMEM_ZEROINIT = &H40
Private Const GMEM_MODIFY = &H80
Private Const GMEM_DISCARDABLE = &H100
Private Const GMEM_NOT_BANKED = &H1000
Private Const GMEM_SHARE = &H2000
Private Const GMEM_DDESHARE = &H2000
Private Const GMEM_NOTIFY = &H4000
Private Const GMEM_LOWER = GMEM_NOT_BANKED
Private Const GMEM_VALID_FLAGS = &H7F72
Private Const GMEM_INVALID_HANDLE = &H8000
Private Const GHND = (GMEM_MOVEABLE Or GMEM_ZEROINIT)
Private Const GPTR = (GMEM_FIXED Or GMEM_ZEROINIT)
Private Const FO_COPY = &H2
'API memory functions
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
Private Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Private Declare Function GlobalSize Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As String) As Long
Dim bdata() As Byte
Private Declare Function GetTickCount& Lib "kernel32" ()

Private Sub Command1_Click()
    Call GetClipboard_ByteArray(Me.hWnd, bdata)
    Print "OK"
End Sub

Private Sub Command2_Click()

    Dim hMem As Long, lSize As Long, lPtr As Long
    If OpenClipboard(Me.hWnd) Then
        If IsClipboardFormatAvailable(CF_TEXT) Then
            hMem = GetClipboardData(CF_TEXT)
            If hMem <> 0 Then
                lSize = GlobalSize(hMem)
                'lSize = lstrlen(hMem)
                If lSize <> 0 Then
                    lPtr = GlobalLock(hMem)
                    If lPtr <> 0 Then
                        ReDim bdata(0 To lSize - 1)
                        CopyMemory bdata(0), ByVal lPtr, lSize
                        GlobalUnlock hMem
                        Print "OK"
                    Else
                        Debug.Print "Âê©w°O¾ÐÅé¥¢±Ñ"
                    End If
                Else
                    Debug.Print "¨ú±o¤º®e¤j¤p¥¢±Ñ=>" & lSize
                End If
            Else
                Debug.Print "¨ú±o¤å¦r¥¢±Ñ"
            
            End If
        Else
            Debug.Print "°Å¶KÃ¯¤¤¨S¦³ANSI¤å¦r"
        End If
        CloseClipboard
    Else
        Debug.Print "µLªk¶}±Ò°Å¶KÃ¯"
    End If



End Sub

Private Sub Command3_Click()
    Dim tmpbyte As Byte
    Dim tmpInt As Integer
    tmpInt = Asc("§Ú")
    
    tmpbyte = GetBiAsc(tmpInt, 0)
    '218
    Print OK
    
End Sub

Private Sub Command4_Click()
    Dim hMem As Long, hhMem As Long
    If OpenClipboard(Me.hWnd) Then
        EmptyClipboard
        hMem = GlobalAlloc(GHND, UBound(bdata))
        If hMem Then
            hhMem = GlobalLock(hMem)
            If hhMem Then
                Call CopyMemory(ByVal hhMem, bdata(0), UBound(bdata))
                GlobalUnlock hMem
                SetClipboardData CF_TEXT, hMem
                CloseClipboard
                GlobalFree hMem
            Else
                Debug.Print "Âê©w°O¾ÐÅé¥¢±Ñ"
            End If
            GlobalUnlock hMem
        Else
            Debug.Print "°t¸m°O¾ÐÅé¥¢±Ñ"
        End If
        
        
        Call CloseClipboard
    Else
        Debug.Print "µLªk¶}±Ò°Å¶KÃ¯"
    End If
End Sub

Private Sub Command5_Click()
    Call BA_SetDefault  'ªì©l¤Æbytearray°}¦C¼Ò²Õ¹w³]­È
    Call BA_Reset     'ªì©l¤Æbytearray°}¦C
    Call BA_Put_Str("[m")
End Sub

Private Sub Command6_Click()
    Call TestIf
End Sub

Public Sub TestIf()
    Dim tmpTimeStart As Long, tmpTimeEnd As Long, i As Long
    Dim tmpInt As Integer
    tmpInt = 4
    tmpTimeStart = GetTickCount&()
    For i = 1 To 100000
        If tmpInt >= 32 And tmpInt <= 127 Then
            DoEvents
        ElseIf tmpInt = 10 Then
            DoEvents
        ElseIf tmpInt = 27 Then
            DoEvents
        ElseIf tmpInt > 127 Then
            DoEvents
        Else
            DoEvents
        End If
    Next i
    tmpTimeEnd = GetTickCount&()
    Debug.Print "If:" & (tmpTimeEnd - tmpTimeStart)
End Sub
Public Sub TestSelect()
    Dim tmpTimeStart As Long, tmpTimeEnd As Long, i As Long
    Dim tmpInt As Integer
    tmpInt = 125
    tmpTimeStart = GetTickCount&()
    For i = 1 To 100000
        Select Case tmpInt
            Case 32 To 127
                DoEvents
            Case 10
                DoEvents
            Case 27
                DoEvents
            Case Is > 127
                DoEvents
            Case Else
                DoEvents
        End Select
    Next i
    tmpTimeEnd = GetTickCount&()
    Debug.Print "Select:" & (tmpTimeEnd - tmpTimeStart)
End Sub

Private Sub Command7_Click()
    Call TestSelect
End Sub
