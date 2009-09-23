VERSION 5.00
Begin VB.Form Form20 
   Caption         =   "編譯"
   ClientHeight    =   3585
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7215
   LinkTopic       =   "Form20"
   ScaleHeight     =   3585
   ScaleWidth      =   7215
   StartUpPosition =   3  '系統預設值
End
Attribute VB_Name = "Form20"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'createwindow  "RichEdit"  "RICHEDIT_CLASS"
'RichEdit20A
Const WS_EX_STATICEDGE = &H20000
Const WS_EX_TRANSPARENT = &H20&
Const WS_EX_CLIENTEDGE = &H200
Const WS_CHILD = &H40000000
Const CW_USEDEFAULT = &H80000000
Const SW_NORMAL = 1
Const WS_VISIBLE = &H10000000
Const WS_TABSTOP = &H100000
Const WS_HSCROLL = &H100000
Const WS_VSCROLL = &H200000
Const ES_MULTILINE = &H4&
Const ES_WANTRETURN = &H1000&
'API memory 常數
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
Private Type CREATESTRUCT
    lpCreateParams As Long
    hInstance As Long
    hMenu As Long
    hWndParent As Long
    cy As Long
    cx As Long
    Y As Long
    X As Long
    style As Long
    lpszName As String
    lpszClass As String
    ExStyle As Long
End Type
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hwnd As Long, ByRef lpString As Byte) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Any) As Long
Private Declare Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal nHeight As Long, ByVal nWidth As Long, ByVal nEscapement As Long, ByVal nOrientation As Long, ByVal fnWeight As Long, ByVal fdwItalic As Boolean, ByVal fdwUnderline As Boolean, ByVal fdwStrikeOut As Boolean, ByVal fdwCharSet As Long, ByVal fdwOutputPrecision As Long, ByVal fdwClipPrecision As Long, ByVal fdwQuality As Long, ByVal fdwPitchAndFamily As Long, ByVal lpszFace As String) As Long

Const WM_SETFONT As Integer = &H30
Dim mWnd As Long
Dim hDll As Long
Private Sub Form_Load()
    'KPD-Team 1999
    'URL: http://www.allapi.net/
    'E-Mail: KPDTeam@Allapi.net
    'hDll = LoadLibrary("RichEd32.dll")
    Dim CS As CREATESTRUCT
    Dim m_hFont As Long
    'Create a new label
    'mWnd = CreateWindowEx(WS_CHILD Or WS_VISIBLE Or WS_HSCROLL Or WS_VSCROLL Or ES_MULTILINE Or ES_WANTRETURN, "RichEdit", "Hello World !", WS_CHILD, 0, 0, 300, 50, Me.hwnd, 0, App.hInstance, CS)
    mWnd = CreateWindowEx(WS_EX_CLIENTEDGE, "EDIT", "Hello World !", WS_CHILD Or WS_VISIBLE Or WS_HSCROLL Or WS_VSCROLL Or ES_MULTILINE Or ES_WANTRETURN, 0, 0, Me.Width / 15 - 30, Me.Height / 15 - 90, Me.hwnd, 0, App.hInstance, CS)
    m_hFont = CreateFont(13, 0, 1, 0, 0, 0, 0, 0, 0, 0, 0, 2, 0, "Times New Roman")
    'SendMessage mWnd, WM_SETFONT, m_hFont, ""
    Me.Caption = mWnd
    'Show our label
    ShowWindow mWnd, SW_NORMAL
End Sub
Private Sub Form_Unload(Cancel As Integer)
    'destroy our label
    DestroyWindow mWnd
    'FreeLibrary hDll
End Sub

Public Sub SetText_byteArray()
    'Dim hhmem As Long
    'hMem = GlobalAlloc(GHND, UBound(ByteArray) + 1)
    
    'hhmem = GlobalLock(hMem)
    'Call CopyMemory(ByVal hhmem, ByteArray(0), UBound(ByteArray) + 1)
    
    'GlobalUnlock hMem
    'Call SetWindowText(mWnd, ByVal hMem)
    Call SetWindowText(mWnd, ByteArray(0))
    'SetWindowText
    'GlobalFree hMem

End Sub
