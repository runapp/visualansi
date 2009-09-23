Attribute VB_Name = "ClipboradModule"
Private Declare Function CloseClipboard Lib "user32" () As Long
Private Declare Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function EmptyClipboard Lib "user32" () As Long
Private Declare Function GetClipboardData Lib "user32" (ByVal wFormat As Long) As Long
Private Declare Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
Private Declare Function IsClipboardFormatAvailable Lib "user32" (ByVal wFormat As Long) As Long
Private Enum EPredefinedClipboardFormatConstants
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
'API memory functions
Public Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Public Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
'Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
Private Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Private Declare Function GlobalSize Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As String) As Long
Dim bdata() As Byte
Public Sub GetClipboard_ByteArray(ByVal hwnd As Long, ByRef bArray() As Byte)

    Dim hMem As Long, lSize As Long, lPtr As Long, hAv As Long, hOpen As Long
    hOpen = OpenClipboard(hwnd)
    If hOpen = 0 Then
        Call CloseClipboard
        hOpen = OpenClipboard(hwnd)
    End If
    If hOpen Then
        hAv = IsClipboardFormatAvailable(CF_TEXT)
        If IsClipboardFormatAvailable(CF_TEXT) <> 0 Then
            hMem = GetClipboardData(CF_TEXT)
            If hMem <> 0 Then
                lSize = GlobalSize(hMem)
                'lSize = lstrlen(hMem)
                If lSize <> 0 Then
                    lPtr = GlobalLock(hMem)
                    If lPtr <> 0 Then
                        ReDim bArray(0 To lSize - 1)
                        CopyMemory bArray(0), ByVal lPtr, lSize
                        GlobalUnlock hMem
                        'Print "OK"
                    Else
                        Debug.Print "鎖定記憶體失敗"
                    End If
                Else
                    Debug.Print "取得內容大小失敗=>" & lSize
                End If
            Else
                Debug.Print "取得文字失敗"
            
            End If
        Else
            Debug.Print "剪貼簿中沒有ANSI文字"
        End If
        CloseClipboard
    Else
        Debug.Print "無法開啟剪貼簿"
    End If
End Sub

Public Sub SetClipboard_ByteArray(ByVal hwnd As Long, ByRef bArray() As Byte)
    Dim hMem As Long, hhMem As Long
    hOpen = OpenClipboard(hwnd)
    If hOpen = 0 Then
        Call CloseClipboard
        hOpen = OpenClipboard(hwnd)
    End If
    If hOpen Then
        EmptyClipboard
        hMem = GlobalAlloc(GHND, UBound(bArray) + 1)
        If hMem Then
            hhMem = GlobalLock(hMem)
            If hhMem Then
                Call CopyMemory(ByVal hhMem, bArray(0), UBound(bArray) + 1)
                GlobalUnlock hMem
                SetClipboardData CF_TEXT, hMem
                CloseClipboard
                GlobalFree hMem
            Else
                Debug.Print "鎖定記憶體失敗"
            End If
            GlobalUnlock hMem
        Else
            Debug.Print "配置記憶體失敗"
        End If
        
        Call CloseClipboard
    Else
        Debug.Print "無法開啟剪貼簿"
    End If
End Sub


Public Function GetBiAsc(ByVal ascint As Integer, ByVal RightOrLeft As Byte) As Byte
    Dim tmparr(1) As Byte
    Call CopyMemory(tmparr(0), ascint, 2)
    GetBiAsc = tmparr(RightOrLeft)
    'Debug.Print tmparr(1) & "," & tmparr(0)
End Function
