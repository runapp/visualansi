VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ApiDrawObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Public Traget As Object

Public TwipsPerScaleX As Single
Public TwipsPerScaleY As Single

Private ATraget As PictureBox
Private AhDC As Long
'裏繪圖記憶體
Private hInDC As Long
Private hInBitmap As Long
Private hInObj As Long

Private hBrushColor(16) As Long  '畫筆(填色)調色盤
Private hPenColor(16) As Long    '畫筆(框線)調色盤
Private Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function DrawTextEx Lib "user32" Alias "DrawTextExA" (ByVal hdc As Long, ByVal lpsz As String, ByVal n As Long, lpRect As RECT, ByVal un As Long, ByVal lpDrawTextParams As Any) As Long
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private tmpr As RECT

Private Declare Function CreateDIBSection16 Lib "gdi32" Alias "CreateDIBSection" (ByVal hdc As Long, pBitmapInfo As BitMapInfo16, ByVal un As Long, lplpVoid As Long, ByVal handle As Long, ByVal dw As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long

Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private IP_Out As IPictureDisp  '匯入的圖片(bitmap)
Private IP_outhDC As Long       '匯入圖片之DC的Handel
Private IP_PB As PictureBox     '處理結果的圖片 需要於表單上設立之圖片方塊元件實體

Private IP_Vars(10) As Long '紀錄圖片縮放及位置 0:X座標 1:Y座標 2:寬 3:高 4:源圖寬 5:源圖高 6:源截X 7:源截Y 8:源截寬 9:源截高 10:背景模式(1為啟動)
Private Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type
Private Type BitmapInfoHeader '點陣圖資訊(Win95/NT4與Win98/NT5分別對此結構作了擴充定義以增加功能，在此只用最古老原始定義)
    biSize          As Long     '=40(&H28) 此結構(BitmapInfoHeader)大小
    biWidth         As Long     '點陣圖寬度(單位：像素)
    biHeight        As Long     '點陣圖高度(單位：像素)
    biPlanes        As Integer  '=1
    biBitCount      As Integer  '=1,4,8,24 每個像素以幾個位元儲存
    biCompression   As Long     '=0: 未壓縮, =1: 256 Color RLE8, =2: 16 Color RLE4
    biSizeImage     As Long     '=biWidth * biHeight * biBitCount
    biXPelsPerMeter As Long     '=pixels per meter
    biYPelsPerMeter As Long     '=pixels per meter
    biClrUsed       As Long
    biClrImportant  As Long
End Type
Private Type BitMapInfo16
     bmiHeader As BitmapInfoHeader
     bmiColors(0 To 15) As Long
End Type

'處理半透明
Private Const AC_SRC_OVER = &H0
Private Type BLENDFUNCTION
  BlendOp As Byte
  BlendFlags As Byte
  SourceConstantAlpha As Byte
  AlphaFormat As Byte
End Type
Private Declare Function AlphaBlend Lib "msimg32.dll" (ByVal hdc As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal hdc As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal BLENDFUNCT As Long) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32.dll" (Destination As Any, Source As Any, ByVal Length As Long)

Private Sub Class_Initialize()
    For i = 0 To 15
        hBrushColor(i) = CreateSolidBrush(QBColor(i))
        hPenColor(i) = CreatePen(0, 0, QBColor(i))
    Next i
    
End Sub

Private Sub Class_Terminate()
    For i = 0 To 15
        DeleteObject hBrushColor(i)
        DeleteObject hPenColor(i)
    Next i
    If hInObj <> 0 Then DeleteObject hInObj
End Sub
Public Property Get Traget() As Variant
    If ATraget Is Nothing Then
        Traget = ""
    Else
        Traget = ATraget.Name
    End If
End Property

Public Property Let Traget(ByRef vNewValue As Variant)
    Dim preScaleX As Single
    Dim preScaleY As Single
    Dim preScaleM As Single
    Set ATraget = vNewValue
    AhDC = ATraget.hdc
    preScaleX = ATraget.ScaleWidth
    preScaleY = ATraget.ScaleHeight
    preScaleM = ATraget.ScaleMode
    ATraget.ScaleMode = 3
    
    TwipsPerScaleX = (ATraget.ScaleWidth) / preScaleX
    TwipsPerScaleY = (ATraget.ScaleHeight) / preScaleY
    ATraget.ScaleMode = preScaleM
    ATraget.ScaleWidth = preScaleX
    ATraget.ScaleHeight = preScaleY
    

End Property
Public Sub resetInDC()
    'hInDC
    If hInDC <> 0 Then  '如果原來的dc存在就殺了他
        DeleteObject hInDC
    End If
    hInDC = CreateCompatibleDC(ATraget.hdc)
    '計算寬度
    Dim preScaleX As Single
    Dim preScaleY As Single
    Dim preScaleM As Single
    'Set ATraget = vNewValue
    'AhDC = ATraget.hdc
    preScaleX = ATraget.ScaleWidth
    preScaleY = ATraget.ScaleHeight
    preScaleM = ATraget.ScaleMode
    ATraget.ScaleMode = 3
    
    hInBitmap = CreateCompatibleBitmap(ATraget.hdc, ATraget.ScaleWidth, ATraget.ScaleHeight)
    '回復寬高單位
    ATraget.ScaleMode = preScaleM
    ATraget.ScaleWidth = preScaleX
    ATraget.ScaleHeight = preScaleY
    hInObj = SelectObject(hInDC, hInBitmap)
    If hInObj <> 0 Then DeleteObject hInObj
    
End Sub

Public Sub DrawRectangle(ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal QBcolorcode As Byte)
    SelectObject AhDC, hPenColor(QBcolorcode)
    SelectObject AhDC, hBrushColor(QBcolorcode)
    Rectangle AhDC, X1 * TwipsPerScaleX, Y1 * TwipsPerScaleY, (X2 + 1) * TwipsPerScaleX, (Y2 + 1) * TwipsPerScaleY
    '裏畫布
    'SelectObject hInDC, hPenColor(QBcolorcode)
    'SelectObject hInDC, hBrushColor(QBcolorcode)
    'Rectangle hInDC, x1 * TwipsPerScaleX, Y1 * TwipsPerScaleY, (X2 + 1) * TwipsPerScaleX, (Y2 + 1) * TwipsPerScaleY
    
End Sub


Public Sub SetTraget()

    AhDC = ATraget.hdc

End Sub

Public Sub PrintText(ByVal text As String, ByVal Color As Long, ByVal X As Long, ByVal Y As Long)
    SetTextColor ATraget.hdc, Color

    X = X * TwipsPerScaleX
    Y = Y * TwipsPerScaleY
    
    'TextOut hInDC, X, Y, text, LenB(text)
    
    TextOut ATraget.hdc, X, Y, text, LenB(text)

End Sub

Public Sub PrintText_biByte_Left(ByVal text As String, ByVal Color As Long, ByVal X As Long, ByVal Y As Long)
    'hInDC
    'SetTextColor ATraget.hdc, Color
    SetTextColor ATraget.hdc, Color
    
    X = X * TwipsPerScaleX
    Y = Y * TwipsPerScaleY
    tmpr.Left = X
    tmpr.Right = X + TwipsPerScaleX
    tmpr.Top = Y
    tmpr.Bottom = Y + TwipsPerScaleY
    'TextOut ATraget.hDC, X, Y, text, 1
    DrawTextEx ATraget.hdc, text, LenB(text), tmpr, ByVal 0&, ByVal 0&
    'DrawTextEx hInDC, text, LenB(text), tmpR, ByVal 0&, ByVal 0&
'ATraget.Refresh
End Sub
Public Sub resetBG(ByVal X As Long, ByVal Y As Long)
    On Error Resume Next
    If X <= UBound(Arrf, 1) Then
        If Arrf(X + 1, Y, OFP.CurrentPage).Ansi = -1 Then
            
            If IP_Vars(10) = 1 Then
                Call ReShow_BG(X, Y, X + 2, Y)
            Else
                DrawRectangle X, Y, X + 1, Y, 0
            End If
        ElseIf Arrf(X, Y, OFP.CurrentPage).Ansi = -1 Then
            
            If IP_Vars(10) = 1 Then
                Call ReShow_BG(X - 1, Y, X + 1, Y)
            Else
                DrawRectangle X - 1, Y, X, Y, 0
            End If
        Else
            If IP_Vars(10) = 1 Then
                Call ReShow_BG(X, Y, X + 1, Y)
            Else
                DrawRectangle X, Y, X, Y, 0
            End If
        End If
    Else
        If IP_Vars(10) = 1 Then
            Call ReShow_BG(X, Y, X + 1, Y)
        Else
            DrawRectangle X, Y, X, Y, 0
        End If
    End If
End Sub
Public Sub resetBG_Area(ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long)
    For i = X1 To X2
        For j = Y1 To Y2
        
            Call resetBG(i, j)
        Next j
    
    Next i
End Sub

Public Sub ShowIt(ByVal X As Single, ByVal Y As Single)
On Error GoTo out
    '新版本的顯示函式
    '支援雙色字
    
    If X < UBound(Arrf, 1) Then
        
        If Arrf(X + 1, Y, OFP.CurrentPage).Ansi = -1 Then
            DrawRectangle X, Y, X, Y, Arrf(X, Y, OFP.CurrentPage).BColor
            DrawRectangle X + 1, Y, X + 1, Y, Arrf(X + 1, Y, OFP.CurrentPage).BColor
            PrintText Chr(Arrf(X, Y, OFP.CurrentPage).Ansi), QBColor(Arrf(X + 1, Y, OFP.CurrentPage).Color), X, Y
            PrintText_biByte_Left Chr(Arrf(X, Y, OFP.CurrentPage).Ansi), QBColor(Arrf(X, Y, OFP.CurrentPage).Color), X, Y
        ElseIf Arrf(X, Y, OFP.CurrentPage).Ansi = -1 Then
            DrawRectangle X - 1, Y, X - 1, Y, Arrf(X - 1, Y, OFP.CurrentPage).BColor
            DrawRectangle X, Y, X, Y, Arrf(X, Y, OFP.CurrentPage).BColor
            PrintText Chr(Arrf(X - 1, Y, OFP.CurrentPage).Ansi), QBColor(Arrf(X, Y, OFP.CurrentPage).Color), X - 1, Y
            PrintText_biByte_Left Chr(Arrf(X - 1, Y, OFP.CurrentPage).Ansi), QBColor(Arrf(X - 1, Y, OFP.CurrentPage).Color), X - 1, Y
        ElseIf Arrf(X, Y, OFP.CurrentPage).BColor = 0 And (Arrf(X, Y, OFP.CurrentPage).Ansi = 0 Or Arrf(X, Y, OFP.CurrentPage).Ansi = 32) Then
            If IP_Vars(10) = 1 Then
                Call ReShow_BG(X, Y, X, Y)
            Else
                DrawRectangle X, Y, X, Y, 0
            End If
        Else
            DrawRectangle X, Y, X, Y, Arrf(X, Y, OFP.CurrentPage).BColor
            PrintText Chr(Arrf(X, Y, OFP.CurrentPage).Ansi), QBColor(Arrf(X, Y, OFP.CurrentPage).Color), X, Y
        End If
    Else
        If Arrf(X, Y, OFP.CurrentPage).Ansi = -1 Then
            DrawRectangle X - 1, Y, X - 1, Y, Arrf(X - 1, Y, OFP.CurrentPage).BColor
            DrawRectangle X, Y, X, Y, Arrf(X, Y, OFP.CurrentPage).BColor
            PrintText Chr(Arrf(X - 1, Y, OFP.CurrentPage).Ansi), QBColor(Arrf(X, Y, OFP.CurrentPage).Color), X - 1, Y
            PrintText_biByte_Left Chr(Arrf(X - 1, Y, OFP.CurrentPage).Ansi), QBColor(Arrf(X - 1, Y, OFP.CurrentPage).Color), X - 1, Y
        ElseIf Arrf(X, Y, OFP.CurrentPage).BColor = 0 And (Arrf(X, Y, OFP.CurrentPage).Ansi = 0 Or Arrf(X, Y, OFP.CurrentPage).Ansi = 32) Then
            If IP_Vars(10) = 1 Then
                Call ReShow_BG(X, Y, X, Y)
            Else
                DrawRectangle X, Y, X, Y, 0
            End If
        Else
            DrawRectangle X, Y, X, Y, Arrf(X, Y, OFP.CurrentPage).BColor
            PrintText Chr(Arrf(X, Y, OFP.CurrentPage).Ansi), QBColor(Arrf(X, Y, OFP.CurrentPage).Color), X, Y
        End If
    End If
    
Exit Sub
out:
    Debug.Print "ShowIt::Error:" & Err.Description & "->In:" & "(" & X & "," & Y & ")"
End Sub
Public Sub ReDraw()
    
    Dim tmpX As Long, tmpY As Long
    'ATraget.Cls
   ' ATraget.AutoRedraw = False
    Call SetTraget
    tmpX = UBound(Arrf, 1)
    tmpY = UBound(Arrf, 2)
    If IP_Vars(10) = 1 Then Call ReShow_BG(0, 0, tmpX, tmpY)
    For i = 0 To tmpY
        For j = 0 To tmpX
            Call ShowIt(j, i)
            DoEvents
        Next j
    Next i
    
    'StretchBlt hInDC, 0, 0, tmpX * TwipsPerScaleX, tmpY * TwipsPerScaleY, ATraget.hdc, 0, 0, tmpX * TwipsPerScaleX, tmpY * TwipsPerScaleY, vbSrcCopy
    ATraget.Refresh
    'ATraget.AutoRedraw = True
End Sub
Public Sub ReDraw_Area(ByVal X1 As Integer, ByVal Y1 As Integer, ByVal X2 As Integer, ByVal Y2 As Integer)
    'If IP_Vars(10) = 1 Then Call ReShow_BG(X1, Y1, X2, Y2)
    For i = Y1 To Y2
        For j = X1 To X2
            Call ShowIt(j, i)
            DoEvents
        Next j
        'DoEvents
    Next i
    
    ATraget.Refresh
End Sub

'處理圖片
Public Sub LoadIP(ByVal FilePath As String)
    On Error GoTo out
    Set IP_Out = LoadPicture(FilePath)
    IP_Vars(0) = 0
    IP_Vars(1) = 0
    Call GetBitmapWH(IP_Out.handle)
    IP_Vars(2) = IP_Vars(4)
    IP_Vars(3) = IP_Vars(5)
    IP_Vars(6) = 0
    IP_Vars(7) = 0
    IP_Vars(8) = IP_Vars(4)
    IP_Vars(9) = IP_Vars(5)
    
    Call Set_IP_outhDC
    
    'IP_Vars(4) = 0
    'Call DrawBitmap2DC(hDC, IP_Out.Handle, 0, 0)
    'Call ReShow_BG(0, 0, 5, 5)
    Exit Sub
out:
    Debug.Print "LoadIP::Err:" & Err.Description
End Sub

Private Sub Set_IP_outhDC()
    '設定來源圖的Dc
    '
    If IP_outhDC <> 0 Then  '如果原來的dc存在就殺了他
        DeleteObject IP_outhDC
        
    End If
    IP_outhDC = CreateCompatibleDC(ATraget.hdc) '建立符合主畫布的dc
    'Debug.Print "建立dc"
    SelectObject IP_outhDC, IP_Out.handle
    'Debug.Print "dc選擇bitmap:" & IP_outhDC & "<-SEL->" & IP_Out.Handle
End Sub
Public Sub GetBitmapWH(ByVal hbitmap As Long)
    Dim bm As BITMAP
    GetObject hbitmap, Len(bm), bm
    IP_Vars(4) = bm.bmWidth
    IP_Vars(5) = bm.bmHeight
    Debug.Print "BITMAP INFO " & bm.bmWidthBytes & ";" & bm.bmPlanes & ";" & bm.bmBitsPixel
End Sub
Public Sub Load_IP_PB(ByRef PB As PictureBox)
    Set IP_PB = PB
End Sub

Public Function GetIP_Vars(ByVal Index As Byte) As Integer
    '紀錄圖片縮放及位置 0:X座標 1:Y座標 2:寬 3:高 ......詳見宣告說明
    GetIP_Vars = IP_Vars(Index)
End Function

Public Sub ShowIPout2PB()
    'Dim BF As BLENDFUNCTION, lBF As Long
    IP_PB.Cls
    StretchBlt IP_PB.hdc, IP_Vars(0), IP_Vars(1), IP_Vars(2), IP_Vars(3), IP_outhDC, IP_Vars(6), IP_Vars(7), IP_Vars(8), IP_Vars(9), vbSrcCopy
   ' With BF
    '   .BlendOp = AC_SRC_OVER
    '    .BlendFlags = 0
    '    .SourceConstantAlpha = 128
    '    .AlphaFormat = 0
    'End With
    'copy the BLENDFUNCTION-structure to a Long
    'RtlMoveMemory lBF, BF, 4
    'lBF = vbBlue - 128 * (vbYellow + 1)
    'AlphaBlend the picture from Picture1 over the picture of Picture2
    'AlphaBlend ATraget.hdc, IP_Vars(0), IP_Vars(1), IP_Vars(2), IP_Vars(3), IP_outhDC, IP_Vars(6), IP_Vars(7), IP_Vars(8), IP_Vars(9), lBF

    
    IP_PB.Refresh
End Sub
Private Sub DrawBitmap2DC(ByVal thdc As Long, ByVal hbitmap As Long, ByVal X As Integer, ByVal Y As Integer)
    Dim bm As BITMAP, tmphDC As Long
    tmphDC = CreateCompatibleDC(thdc)  '建立相容的於目標dc的dc
    SelectObject tmphDC, hbitmap
    GetObject hbitmap, Len(bm), bm
    'Debug.Print "bitmap info " & bm.bmWidth & ":" & bm.bmHeight
    BitBlt thdc, X, Y, bm.bmWidth, bm.bmHeight, tmphDC, 0, 0, vbSrcCopy
    'BitBlt thdc, X, Y, 50, 50, tmphDC, 0, 0, vbSrcCopy
    DeleteDC tmphDC
End Sub
Public Sub ReShow_BG(ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long)
    On Error GoTo out
    '改良版的顯示背景
    'Dim tmphDC As Long
    Dim BF As BLENDFUNCTION, lBF As Long, W As Long, H As Long
    X1 = X1 * TwipsPerScaleX
    Y1 = Y1 * TwipsPerScaleY
    X2 = X2 * TwipsPerScaleX
    Y2 = Y2 * TwipsPerScaleY
    W = X2 - X1 + TwipsPerScaleX
    H = Y2 - Y1 + TwipsPerScaleY

    With BF
        .BlendOp = AC_SRC_OVER
        .BlendFlags = 0
        .SourceConstantAlpha = 128
        .AlphaFormat = 0
    End With
    'copy the BLENDFUNCTION-structure to a Long
    RtlMoveMemory lBF, BF, 4
    IP_PB.Refresh
    'lBF = vbBlue - 128 * (vbYellow + 1)
    'AlphaBlend the picture from Picture1 over the picture of Picture2
    'AlphaBlend ATraget.hdc, x1, Y1, X2 - x1 + TwipsPerScaleX, Y2 - Y1 + TwipsPerScaleY, IP_outhDC, x1, Y1, X2 - x1 + TwipsPerScaleX, Y2 - Y1 + TwipsPerScaleY, lBF
    'AlphaBlend ATraget.hdc, IP_Vars(0), IP_Vars(1), IP_Vars(2), IP_Vars(3), IP_outhDC, IP_Vars(6), IP_Vars(7), IP_Vars(8), IP_Vars(9), lBF
    
    BitBlt ATraget.hdc, X1, Y1, W, H, IP_PB.hdc, X1, Y1, vbSrcCopy
    'Debug.Print "BitBlt:"; &HDC & "," & x1 & "," & Y1 & "," & X2 - x1 & "," & Y2 - Y1 & "," & tmphDC & "," & x1 & "," & Y1 & ","
    'DeleteDC tmphDC
    Exit Sub
out:
    Debug.Print "ReShow_BG::Error:" & Err.Description
End Sub
    'IPED系列函數 -- 調整圖片的函數
    
Public Sub IPED_SetVars(ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long)
    IP_Vars(0) = X
    IP_Vars(1) = Y
    IP_Vars(2) = Width
    IP_Vars(3) = Height
    Call ShowIPout2PB
    
End Sub
Public Sub IPED_FitPBSize()
    IP_Vars(0) = 0
    IP_Vars(1) = 0
    Dim tmpX As Integer, tmpY As Integer
    tmpX = UBound(Arrf, 1)
    tmpY = UBound(Arrf, 2)
    If (tmpX + 1) * IP_Vars(9) < (tmpY + 1) * IP_Vars(8) Then
        '以寬為準
        'IP_Vars(2) = (VAFileH.ArrXUbound + 1) * PPCL
        'OFP
        IP_Vars(2) = (tmpX + 1) * TwipsPerScaleX
        Debug.Print (IP_Vars(2) * IP_Vars(9)) / IP_Vars(8)
        IP_Vars(3) = Fix((IP_Vars(2) * IP_Vars(9)) / IP_Vars(8))

    Else
        '以高為準
        IP_Vars(3) = (tmpY + 1) * TwipsPerScaleY
        IP_Vars(2) = Fix((IP_Vars(3) * IP_Vars(8)) / IP_Vars(9))

    End If
    Call ShowIPout2PB
End Sub
Public Sub IPED_SizePixelCube()
    '將像素放大到跟方格一樣大
    'IP_Vars(0) = IP_Vars(0) * PPCL
    'IP_Vars(1) = IP_Vars(1) * PPCL
    IP_Vars(2) = IP_Vars(8) * 14
    IP_Vars(3) = IP_Vars(9) * 14

    Call ShowIPout2PB
End Sub
Public Sub IPED_Restore()
'還原圖片原貌
    IP_Vars(0) = 0
    IP_Vars(1) = 0
    IP_Vars(2) = IP_Vars(4)
    IP_Vars(3) = IP_Vars(5)
    IP_Vars(6) = 0
    IP_Vars(7) = 0
    IP_Vars(8) = IP_Vars(4)
    IP_Vars(9) = IP_Vars(5)
    Call ShowIPout2PB
End Sub
Public Sub IPED_Cut(ByVal X As Long, ByVal Y As Long, ByVal W As Long, ByVal H As Long)
'剪裁
'紀錄圖片縮放及位置 0:X座標 1:Y座標 2:寬 3:高 4:源圖寬 5:源圖高 6:源截X 7:源截Y 8:源截寬 9:源截高 10:背景模式(1為啟動)
    Dim Rate_X As Double, Rate_Y As Double
    'W = W * PPCL
    'H = H * PPCL
    Rate_X = IP_Vars(8) / IP_Vars(2)
    Rate_Y = IP_Vars(9) / IP_Vars(3)
    Debug.Print "Rate_X=" & Rate_X & ";Rate_Y" & Rate_Y
    IP_Vars(6) = Rate_X * X + IP_Vars(6) - IP_Vars(0) * Rate_X
    IP_Vars(7) = Rate_Y * Y + IP_Vars(7) - IP_Vars(1) * Rate_Y
    
    IP_Vars(8) = Rate_X * W
    IP_Vars(9) = Rate_Y * H
    
    IP_Vars(0) = X
    IP_Vars(1) = Y
    IP_Vars(2) = W
    IP_Vars(3) = H
    'For i = 0 To 9
    '    tmpStr1 = tmpStr1 & "(" & i & ")=" & IP_Vars(i)
    'Next i
    'Debug.Print tmpStr1
    Call ShowIPout2PB
    
End Sub
Public Sub IPED_TransForm(ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long)
    '變形
    IP_Vars(0) = X1
    IP_Vars(1) = Y1
    IP_Vars(2) = X2 - X1 + 1
    IP_Vars(3) = Y2 - Y1 + 1
    Call ShowIPout2PB
End Sub
Public Sub IPED_SetTopLeft(ByVal X As Long, ByVal Y As Long)
    IP_Vars(0) = IP_Vars(0) - X
    IP_Vars(1) = IP_Vars(1) - Y
    Call ShowIPout2PB
End Sub
Public Sub IPED_SetPos(ByVal X As Long, ByVal Y As Long)
    IP_Vars(0) = X
    IP_Vars(1) = Y
    Call ShowIPout2PB
End Sub
'是否使用背景
Public Property Get HaveBG() As Variant
    HaveBG = IP_Vars(10)
End Property

Public Property Let HaveBG(ByVal vNewValue As Variant)
    IP_Vars(10) = vNewValue
End Property
