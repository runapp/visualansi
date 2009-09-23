VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CubeDrawerObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'依賴module1
Private Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal x1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
'Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long '已在module1背宣告
'外部輸入圖片召換式
'宣告bitmap資料格式
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
Private Declare Function CreateDIBSection16 Lib "gdi32" Alias "CreateDIBSection" (ByVal hdc As Long, pBitmapInfo As BitMapInfo16, ByVal un As Long, lplpVoid As Long, ByVal handle As Long, ByVal dw As Long) As Long

Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Const DIB_RGB_COLORS = 0&
'程式自行調用變數
Private hBrushColor(8) As Long  '畫筆(填色)調色盤
Private hPenColor(8) As Long    '畫筆(框線)調色盤
Private hdc As Long             '主要顯示介面之hDC
Private PB As PictureBox        '主要顯示介面
Private MainArr() As Byte       '主記憶體畫布   需要於表單上設立之圖片方塊元件實體
Public CurrentColor As Byte     '目前選取色彩
Private cPage As Integer        '目前所在頁數
Private tmpbyte As Byte         'byte公用暫存
Private VAFileH As VAFileInfo   '檔案資訊
Public cFilepath As String
Private CPBArr() As Byte        '剪貼簿之暫存
Public CPB_uX As Integer       '紀錄剪貼簿之變數uboundX
Public CPB_uY As Integer       '紀錄剪貼簿之變數uboundY
Private IP_Out As IPictureDisp  '匯入的圖片(bitmap)
Private IP_outhDC As Long       '匯入圖片之DC的Handel
Private IP_PB As PictureBox     '處理結果的圖片 需要於表單上設立之圖片方塊元件實體
'Private IP_inhBM As Long        '處理結果的圖片的hbitmap
'Private IP_inhDC As Long        '處理結果的圖片的hdc

Private IP_Vars(10) As Long   '紀錄圖片縮放及位置 0:X座標 1:Y座標 2:寬 3:高 4:源圖寬 5:源圖高 6:源截X 7:源截Y 8:源截寬 9:源截高 10:背景模式(1為啟動)
Private QBLC(15) As Long  '紀錄qbcolor(0~15)的long值
Private PPCL As Byte     'Pixels Per Cube Len 每個方格長度所含之像素
'色彩分析系列
Private CA_Vars(7) As Integer   '0~2 R G B暫存 3~5 R G B閥值  6最暗色  7最亮色
'記憶體畫布的格式
'寬w 高h 頁數z
'arr(0 to (w-1)+1,0 to (h-1),1 to z) as byte
'內各元素儲存0~7之qbcolor code
'此格式雖然很浪費空間但是為了程式設計方便所需不得有所犧牲

'Private ByteArray() As Byte  '以位元組陣列的方式 儲存輸出的彩色碼 已於ByteArrayMudule宣告


Private Sub Class_Initialize()

    '建立GDI調色盤
    For i = 0 To 7
        hBrushColor(i) = CreateSolidBrush(QBColor(i))
        hPenColor(i) = CreatePen(0, 0, QBColor(i))
    Next i
    For i = 0 To 15
        QBLC(i) = QBColor(i)
    Next i
    '設定方塊大小
    PPCL = 10
    '設定自製色彩分析函數亮暗界
    Call CA_v2_BC_DEFUALT
End Sub

Private Sub Class_Terminate()
    '還原設定
    SelectObject PB.hdc, hBrushColor(8)
    SelectObject PB.hdc, hPenColor(8)
    DeleteObject IP_outhDC  '
    '移除GDI調色盤
    For i = 0 To 8
        DeleteObject hBrushColor(i)
        DeleteObject hPenColor(i)
    Next i
End Sub

Public Property Get H() As Variant
    H = VAFileH.ArrYUbound + 1
End Property

Public Property Get W() As Variant
    W = VAFileH.ArrXUbound + 1
End Property
Public Property Get Z() As Variant
    Z = VAFileH.ArrZLenth
End Property

Public Property Get TargetPB() As Variant
    TargetPB = PB.Name
End Property

Public Property Let TargetPB(ByVal vNewValue As Variant)
    Set PB = vNewValue
    'hdc = PB.hdc
    '保存預設(此動作尚未確定是否必要)
    hPenColor(8) = SelectObject(PB.hdc, hPenColor(7))
    hBrushColor(8) = SelectObject(PB.hdc, hBrushColor(7))
    PB.ScaleMode = 3
    tmpsw = PB.ScaleWidth
    tmpsh = PB.ScaleHeight
    PB.ScaleMode = 0
    PB.ScaleWidth = tmpsw / PPCL
    PB.ScaleHeight = tmpsh / PPCL
    
End Property

Public Property Get CurrentPage() As Variant
    CurrentPage = cPage
End Property

Public Property Let CurrentPage(ByVal vNewValue As Variant)
    If vNewValue <= VAFileH.ArrZLenth And vNewValue >= 1 Then    '排除不合理的頁數要求
        cPage = vNewValue
        Call ReShow(0, 0, VAFileH.ArrXUbound, VAFileH.ArrYUbound)   '更新畫面
    End If
End Property
Public Sub SetColor(ByVal QBcolorcode As Byte)
    On Error GoTo out
    '設定顏色
    SelectObject PB.hdc, hPenColor(QBcolorcode)
    SelectObject PB.hdc, hBrushColor(QBcolorcode)
    CurrentColor = QBcolorcode
    Exit Sub
out:
    Debug.Print "SetColor error ;QBcolorcode=" & QBcolorcode
End Sub
Public Sub SetDC(ByVal ihdc As Long)
    hdc = ihdc
    '給dc套用色盤預設值
    hPenColor(8) = SelectObject(hdc, hPenColor(7))
    hBrushColor(8) = SelectObject(hdc, hBrushColor(7))
End Sub
Public Sub ReSetDC()
    'picturebox在clear後dc會變掉 故要重新指定
    hdc = PB.hdc
    Call SetColor(CurrentColor)
End Sub

Public Sub SetSize(ByVal X As Integer, ByVal Y As Integer, ByVal Z As Integer, Optional ByVal dopreserve As Byte)
'X Y Z皆指格數

    If dopreserve = 1 Then
        ReDim Preserve MainArr(X - 1 + 1, Y - 1, 1 To Z)    '(x-1)+1 之+1是為了讓編譯時方便(因有判斷i+1之情境)
        '若畫布縮小則需確定+1的那個部分為空白
        For k = 1 To Z
            For j = 0 To (Y - 1)
                MainArr(X, j, k) = 0
                DoEvents
            Next j
        Next k
    Else
        ReDim MainArr(X - 1 + 1, Y - 1, 1 To Z)             '(x-1)+1 之+1是為了讓編譯時方便(因有判斷i+1之情境)
    End If
    PB.Parent.ScaleMode = 3
    PB.Width = PPCL * X + 5
    PB.Height = PPCL * Y + 5
    VAFileH.ArrXUbound = X - 1
    VAFileH.ArrYUbound = Y - 1
    VAFileH.ArrZLenth = Z
    '設定初始頁
    cPage = 1
    'PB.Cls
    Call ReSetDC
    Call SetColor(CurrentColor)
End Sub
Public Sub DrawDC_Point(ByVal X As Integer, ByVal Y As Integer)
    pixX = X * PPCL
    pixY = Y * PPCL
    Rectangle PB.hdc, pixX, pixY, pixX + PPCL, pixY + PPCL
    
End Sub


Public Sub ReShow(ByVal x1 As Integer, ByVal Y1 As Integer, ByVal X2 As Integer, ByVal Y2 As Integer)
    Dim preColor As Integer
    preColor = CurrentColor
    If IP_Vars(10) = 1 Then '如果啟動背景
        Call ReShow_BG(x1, Y1, X2, Y2)
        For j = Y1 To Y2
            For i = x1 To X2
                If MainArr(i, j, cPage) <> CurrentColor Then
                    Call SetColor(MainArr(i, j, cPage))
                End If
                If MainArr(i, j, cPage) <> 0 Then
                    Call DrawDC_Point(i, j)
                End If
            Next i
        Next j
    Else
        For j = Y1 To Y2
            For i = x1 To X2
                If MainArr(i, j, cPage) <> CurrentColor Then
                    Call SetColor(MainArr(i, j, cPage))
                End If
                Call DrawDC_Point(i, j)
            Next i
        Next j
        
    End If
    PB.Refresh
    Call SetColor(preColor) '還原色彩
End Sub

Public Sub ReShow_Point(ByVal X As Integer, ByVal Y As Integer)
    Dim preColor As Integer
    preColor = CurrentColor
    Call SetColor(MainArr(X, Y, cPage))
    Call DrawDC_Point(X, Y)
    Call SetColor(preColor) '還原色彩
End Sub

'#### 編譯 ####
Public Sub SetByteArray()
    '初始化bytearray模組設定值
    Call BA_SetDefault
    Call BA_Reset   '初始化bytearray
End Sub
Public Sub BA_ClipBoard_Copy(ByVal hwnd As Long)
    Call BA_CutTail
    Call SetClipboard_ByteArray(hwnd, ByteArray)
End Sub
Public Sub BA_SaveAnsiFile(ByVal filepath As String)
    Dim outfile As Integer
    Call BA_CutTail
    outfile = 40
    Open filepath For Binary As #outfile
    Put #outfile, 1, ByteArray
    Close outfile
End Sub
Public Function GetAnsiText_All(Optional deColorLost As Byte, Optional addPageNum As Byte, Optional FixBlankEnd As Byte)
    Dim rv As Variant     '回傳暫存
    Dim PageLine As String '頁與頁之間的插入行(為了動畫設計)
    Dim tmpInt(3) As Integer
    tmpInt(0) = 0
    tmpInt(1) = 0
    tmpInt(2) = VAFileH.ArrXUbound
    tmpInt(3) = VAFileH.ArrYUbound
    If VAFileH.filetype = 4 Then
        PageLine = ""
    Else
        PageLine = "" & "---" & vbCrLf
    End If

        For k = 1 To VAFileH.ArrZLenth
            If addPageNum = 1 Then
                'rv = rv & GetAnsiText_Area(tmpInt(0), tmpInt(1), tmpInt(2), tmpInt(3), k, deColorLost, FixBlankEnd) & "Page." & k & PageLine
                'Call GetAnsiText_Area(tmpInt(0), tmpInt(1), tmpInt(2), tmpInt(3), k, deColorLost, FixBlankEnd)
                Call GetAnsiText_Area(tmpInt(0), tmpInt(1), tmpInt(2), tmpInt(3), k, 1, 1)
                Call BA_Put_Str("Page." & k & PageLine)
            Else
                'rv = rv & GetAnsiText_Area(tmpInt(0), tmpInt(1), tmpInt(2), tmpInt(3), k, deColorLost, FixBlankEnd) & PageLine
                'Call GetAnsiText_Area(tmpInt(0), tmpInt(1), tmpInt(2), tmpInt(3), k, deColorLost, FixBlankEnd)
                Call GetAnsiText_Area(tmpInt(0), tmpInt(1), tmpInt(2), tmpInt(3), k, 1, 1)
                
                Call BA_Put_Str(PageLine)
            End If
        Next k

    GetAnsiText_All = rv
    
End Function
Public Function GetAnsiText_All_bak(Optional deColorLost As Byte, Optional addPageNum As Byte, Optional FixBlankEnd As Byte)
    Dim rv As Variant     '回傳暫存
    Dim PageLine As String '頁與頁之間的插入行(為了動畫設計)
    Dim tmpInt(3) As Integer
    tmpInt(0) = 0
    tmpInt(1) = 0
    tmpInt(2) = VAFileH.ArrXUbound
    tmpInt(3) = VAFileH.ArrYUbound
    If VAFileH.filetype = 4 Then
        PageLine = ""
    Else
        PageLine = "" & "---" & vbCrLf
    End If

        For k = 1 To VAFileH.ArrZLenth
            If addPageNum = 1 Then
                rv = rv & GetAnsiText_Area(tmpInt(0), tmpInt(1), tmpInt(2), tmpInt(3), k, deColorLost, FixBlankEnd) & "Page." & k & PageLine
            Else
                rv = rv & GetAnsiText_Area(tmpInt(0), tmpInt(1), tmpInt(2), tmpInt(3), k, deColorLost, FixBlankEnd) & PageLine
            End If
        Next k

    GetAnsiText_All = rv
    
End Function
Public Function GetAnsiText()
    Dim rv As String    '回傳暫存
    Dim tmpColor(1) As Byte '0為前景 1為背景
    Dim preCode As String
    Dim PageLine As String '頁與頁之間的插入行(為了動畫設計)
    
    preCode = Chr(27) & "["
    
    If VAFileH.filetype = 4 Then
        PageLine = vbCrLf
    Else
        PageLine = vbCrLf & "---" & vbCrLf
    End If
    
    'QBCToAnsiBC
    '預設色彩
    tmpColor(0) = 7 '前景淺灰
    tmpColor(1) = 0 '背景黑
    '進行轉換
    rv = rv & preCode & "m"
    For j = 0 To (VAFileH.ArrYUbound + 1) - 2 Step 2  '一個字元的高度為2
        For i = 0 To VAFileH.ArrXUbound
            If MainArr(i, j, cPage) = MainArr(i, j + 1, cPage) Then
                '是否與之前的背景同色(減肥用)
                If MainArr(i, j, cPage) = tmpColor(1) Then
                    rv = rv & " "
                Else
                    rv = rv & preCode & QBCToAnsiBC(MainArr(i, j, cPage)) & "m "
                    tmpColor(1) = MainArr(i, j, cPage)
                End If
            Else
                '四格一起考慮

                
                If MainArr(i, j, cPage) = MainArr(i + 1, j, cPage) And MainArr(i, j + 1, cPage) = MainArr(i + 1, j + 1, cPage) Then
                    If MainArr(i, j, cPage) = tmpColor(1) Then '上半部同
                        If MainArr(i, j + 1, cPage) = tmpColor(0) Then '下半部同
                            rv = rv & "▄"
                        Else
                            rv = rv & preCode & QBC2AnsiC(MainArr(i, j + 1, cPage)) & "m" & "▄"
                            tmpColor(0) = MainArr(i, j + 1, cPage)
                        End If
                    Else
                        If MainArr(i, j + 1, cPage) = tmpColor(0) Then
                            rv = rv & preCode & QBCToAnsiBC(MainArr(i, j, cPage)) & "m" & "▄"
                        Else
                            rv = rv & preCode & QBC2AnsiC(MainArr(i, j + 1, cPage)) & ";" & QBCToAnsiBC(MainArr(i, j, cPage)) & "m" & "▄"
                            tmpColor(0) = MainArr(i, j + 1, cPage)
                        End If
                        tmpColor(1) = MainArr(i, j, cPage)
                    End If
                Else
                    '先考慮左邊
                    If MainArr(i, j, cPage) = tmpColor(1) Then '上半部同
                        If MainArr(i, j + 1, cPage) = tmpColor(0) Then '下半部同
                            rv = rv & "▄" '▄的前半部
                        Else
                            rv = rv & preCode & QBC2AnsiC(MainArr(i, j + 1, cPage)) & "m" & "▄" & Chr(8)
                            tmpColor(0) = MainArr(i, j + 1, cPage)
                        End If
                    Else
                        If MainArr(i, j + 1, cPage) = tmpColor(0) Then
                            rv = rv & preCode & QBCToAnsiBC(MainArr(i, j, cPage)) & "m" & "▄" & Chr(8)
                        Else
                            rv = rv & preCode & QBC2AnsiC(MainArr(i, j + 1, cPage)) & ";" & QBCToAnsiBC(MainArr(i, j, cPage)) & "m" & "▄" & Chr(8)
                            tmpColor(0) = MainArr(i, j + 1, cPage)
                        End If
                        tmpColor(1) = MainArr(i, j, cPage)
                    End If
                    '考慮右邊
                    If MainArr(i + 1, j, cPage) = tmpColor(1) Then '上半部同
                        If MainArr(i + 1, j + 1, cPage) = tmpColor(0) Then '下半部同(在此不可能發生 已於之前排除)
                            rv = rv & "e" '▄的後半部
                            Debug.Print "發生不可能之可能"
                            
                        Else
                            rv = rv & preCode & QBC2AnsiC(MainArr(i + 1, j + 1, cPage)) & "m" & "e"
                            tmpColor(0) = MainArr(i + 1, j + 1, cPage)
                        End If
                    Else
                        If MainArr(i + 1, j + 1, cPage) = tmpColor(0) Then
                            rv = rv & preCode & QBCToAnsiBC(MainArr(i + 1, j, cPage)) & "m" & "e"
                        Else
                            rv = rv & preCode & QBC2AnsiC(MainArr(i + 1, j + 1, cPage)) & ";" & QBCToAnsiBC(MainArr(i + 1, j, cPage)) & "m" & "e"
                            tmpColor(0) = MainArr(i + 1, j + 1, cPage)
                        End If
                        tmpColor(1) = MainArr(i + 1, j, cPage)
                    End If
                End If
                
                '因考慮了兩個位元 故跳過下一個
                i = i + 1
            End If
            
        Next i
        rv = rv & vbCrLf
    Next j
    rv = rv & preCode & "m"
    GetAnsiText = rv

End Function
Public Function GetAnsiText_Area(ByVal x1 As Integer, ByVal Y1 As Integer, ByVal X2 As Integer, ByVal Y2 As Integer, ByVal page As Integer, Optional deColorLost As Byte, Optional FixBlankEnd As Byte)
    On Error GoTo out
    '存入bytearray的版本
    '去除畫面破碎情形
    Dim rv As String    '回傳暫存
    Dim tmpColor(1) As Byte '0為前景 1為背景
    Dim preCode As String, dCLstr As String
    'Dim PageLine As String '頁與頁之間的插入行(為了動畫設計)
    


    preCode = Chr(27) & "["
    'If deColorLost = 1 Then dCLstr = preCode & "m"
    'BA_Put_Str()
    
'    If VAFileH.filetype = 4 Then
'        PageLine = vbCrLf
'    Else
'        PageLine = vbCrLf & "---" & vbCrLf
'    End If
    
    'QBCToAnsiBC
    '預設色彩
    tmpColor(0) = 7 '前景淺灰
    tmpColor(1) = 0 '背景黑
    '進行轉換
    'rv = rv & preCode & "m"
    Call BA_Put_Str(preCode & "m")
    For j = Y1 To (Y2 + 1) - 2 Step 2  '一個字元的高度為2
        For i = x1 To X2
            If MainArr(i, j, page) = MainArr(i, j + 1, page) Then
                '是否與之前的背景同色(減肥用)
                If MainArr(i, j, page) = tmpColor(1) Then
                    'rv = rv & " "
                    Call BA_Put(32)
                Else
                    'rv = rv & preCode & QBCToAnsiBC(MainArr(i, j, page)) & "m "
                    tmpColor(1) = MainArr(i, j, page)
                    Call BA_Put_Str(preCode & QBCToAnsiBC(tmpColor(1)) & "m ")
                    
                End If
            Else
                '四格一起考慮
                If MainArr(i, j, page) = MainArr(i + 1, j, page) And MainArr(i, j + 1, page) = MainArr(i + 1, j + 1, page) Then
                    If MainArr(i, j, page) = tmpColor(1) Then '上半部同
                        If MainArr(i, j + 1, page) = tmpColor(0) Then '下半部同
                            Call BA_Put_Str("▄")
                        Else
                            tmpColor(0) = MainArr(i, j + 1, page)
                            Call BA_Put_Str(preCode & QBC2AnsiC(tmpColor(0)) & "m" & "▄")
                            
                        End If
                    Else
                        tmpColor(1) = MainArr(i, j, page)
                        If MainArr(i, j + 1, page) = tmpColor(0) Then
                            'rv = rv & preCode & QBCToAnsiBC(MainArr(i, j, page)) & "m" & "▄"
                            Call BA_Put_Str(preCode & QBCToAnsiBC(tmpColor(1)) & "m" & "▄")
                        Else
                            'rv = rv & preCode & QBC2AnsiC(MainArr(i, j + 1, page)) & ";" & QBCToAnsiBC(MainArr(i, j, page)) & "m" & "▄"
                            tmpColor(0) = MainArr(i, j + 1, page)
                            Call BA_Put_Str(preCode & QBC2AnsiC(tmpColor(0)) & ";" & QBCToAnsiBC(tmpColor(1)) & "m" & "▄")
                            
                        End If
                        
                    End If
                Else
                    '先考慮左邊
                    If MainArr(i, j, page) = tmpColor(1) Then '上左半部同前色
                        If MainArr(i, j + 1, page) = tmpColor(0) Then '下左半部同前色
                            'rv = rv & "▄" '▄的前半部
                            'Call BA_Put_Str("▄")
                            Call BA_Put(162)
                        Else
                            'rv = rv & preCode & QBC2AnsiC(MainArr(i, j + 1, page)) & "m" & "▄" & Chr(8)
                            tmpColor(0) = MainArr(i, j + 1, page)
                            Call BA_Put_Str(preCode & QBC2AnsiC(tmpColor(0)) & "m")
                            Call BA_Put(162)
                            
                        End If
                    Else
                        tmpColor(1) = MainArr(i, j, page)
                        If MainArr(i, j + 1, page) = tmpColor(0) Then
                            
                            Call BA_Put_Str(preCode & QBCToAnsiBC(tmpColor(1)) & "m")
                            Call BA_Put(162)
                            
                            'rv = rv & preCode & QBCToAnsiBC(MainArr(i, j, page)) & "m" & "▄" & Chr(8)
                        Else
                            tmpColor(0) = MainArr(i, j + 1, page)
                            Call BA_Put_Str(preCode & QBC2AnsiC(tmpColor(0)) & ";" & QBCToAnsiBC(tmpColor(1)) & "m")
                            Call BA_Put(162)
                            'rv = rv & preCode & QBC2AnsiC(MainArr(i, j + 1, page)) & ";" & QBCToAnsiBC(MainArr(i, j, page)) & "m" & "▄" & Chr(8)
                            
                        End If
                        
                    End If
                    '考慮右邊
                    If MainArr(i + 1, j, page) = tmpColor(1) Then '上半部同
                        If MainArr(i + 1, j + 1, page) = tmpColor(0) Then '下半部同(在此不可能發生 已於之前排除)
                            'rv = rv & "e" '▄的後半部
                            Call BA_Put(101)
                            Debug.Print "發生不可能之可能"
                            
                        Else
                            'rv = rv & preCode & QBC2AnsiC(MainArr(i + 1, j + 1, page)) & "m" & "e"
                            tmpColor(0) = MainArr(i + 1, j + 1, page)
                            Call BA_Put_Str(preCode & QBC2AnsiC(tmpColor(0)) & "m")
                            Call BA_Put(101)
                            
                        End If
                    Else
                        tmpColor(1) = MainArr(i + 1, j, page)
                        If MainArr(i + 1, j + 1, page) = tmpColor(0) Then
                            'rv = rv & preCode & QBCToAnsiBC(MainArr(i + 1, j, page)) & "m" & "e"
                            Call BA_Put_Str(preCode & QBCToAnsiBC(tmpColor(1)) & "m")
                            Call BA_Put(101)
                        Else
                            tmpColor(0) = MainArr(i + 1, j + 1, page)
                            Call BA_Put_Str(preCode & QBC2AnsiC(tmpColor(0)) & ";" & QBCToAnsiBC(tmpColor(1)) & "m")
                            Call BA_Put(101)
                            'rv = rv & preCode & QBC2AnsiC(MainArr(i + 1, j + 1, page)) & ";" & QBCToAnsiBC(MainArr(i + 1, j, page)) & "m" & "e"
                            
                        End If
                        
                    End If
                End If
                
                '因考慮了兩個位元 故跳過下一個
                i = i + 1
            End If
        Next i
        '以下為修正畫面破裂程式碼
        If FixBlankEnd = 1 Then
            If tmpColor(1) <> 0 Then
                'rv = rv & preCode & "m" & vbCrLf
                Call BA_Put_Str(preCode & "m" & vbCrLf)
                tmpColor(0) = 7 '前景淺灰
                tmpColor(1) = 0 '背景黑
            Else
                If deColorLost = 1 Then
                    tmpColor(0) = 7 '前景淺灰
                    tmpColor(1) = 0 '背景黑
                    'rv = rv & vbCrLf & dCLstr
                    Call BA_Put_Str(vbCrLf & preCode & "m")
                Else
                    'rv = rv & vbCrLf
                    Call BA_Put_Str(vbCrLf)
                End If
               
            End If
        Else
            If deColorLost = 1 Then
                tmpColor(0) = 7 '前景淺灰
                tmpColor(1) = 0 '背景黑
                'rv = rv & vbCrLf & preCode & "m"
                Call BA_Put_Str(vbCrLf & preCode & "m")
            Else
                'rv = rv & vbCrLf
                Call BA_Put_Str(vbCrLf)
            End If
            
        End If
    Next j
    BA_Put_Str (preCode & "m")
    'rv = rv & preCode & "m"
    'GetAnsiText_Area = rv
Exit Function
out:
Debug.Print "編譯錯誤:i,j=" & i & "," & j
Resume Next

End Function

Public Function GetAnsiText_Area_bak(ByVal x1 As Integer, ByVal Y1 As Integer, ByVal X2 As Integer, ByVal Y2 As Integer, ByVal page As Integer, Optional deColorLost As Byte, Optional FixBlankEnd As Byte)
    '去除畫面破碎情形
    Dim rv As String    '回傳暫存
    Dim tmpColor(1) As Byte '0為前景 1為背景
    Dim preCode As String, dCLstr As String
    'Dim PageLine As String '頁與頁之間的插入行(為了動畫設計)
    

    preCode = Chr(27) & "["
    If deColorLost = 1 Then dCLstr = preCode & "m"
    
    
'    If VAFileH.filetype = 4 Then
'        PageLine = vbCrLf
'    Else
'        PageLine = vbCrLf & "---" & vbCrLf
'    End If
    
    'QBCToAnsiBC
    '預設色彩
    tmpColor(0) = 7 '前景淺灰
    tmpColor(1) = 0 '背景黑
    '進行轉換
    rv = rv & preCode & "m"
    For j = Y1 To (Y2 + 1) - 2 Step 2  '一個字元的高度為2
        For i = x1 To X2
            If MainArr(i, j, page) = MainArr(i, j + 1, page) Then
                '是否與之前的背景同色(減肥用)
                If MainArr(i, j, page) = tmpColor(1) Then
                    rv = rv & " "
                Else
                    rv = rv & preCode & QBCToAnsiBC(MainArr(i, j, page)) & "m "
                    tmpColor(1) = MainArr(i, j, page)
                End If
            Else
                '四格一起考慮
                If MainArr(i, j, page) = MainArr(i + 1, j, page) And MainArr(i, j + 1, page) = MainArr(i + 1, j + 1, page) Then
                    If MainArr(i, j, page) = tmpColor(1) Then '上半部同
                        If MainArr(i, j + 1, page) = tmpColor(0) Then '下半部同
                            rv = rv & "▄"
                        Else
                            rv = rv & preCode & QBC2AnsiC(MainArr(i, j + 1, page)) & "m" & "▄"
                            tmpColor(0) = MainArr(i, j + 1, page)
                        End If
                    Else
                        If MainArr(i, j + 1, page) = tmpColor(0) Then
                            rv = rv & preCode & QBCToAnsiBC(MainArr(i, j, page)) & "m" & "▄"
                        Else
                            rv = rv & preCode & QBC2AnsiC(MainArr(i, j + 1, page)) & ";" & QBCToAnsiBC(MainArr(i, j, page)) & "m" & "▄"
                            tmpColor(0) = MainArr(i, j + 1, page)
                        End If
                        tmpColor(1) = MainArr(i, j, page)
                    End If
                Else
                    '先考慮左邊
                    If MainArr(i, j, page) = tmpColor(1) Then '上半部同
                        If MainArr(i, j + 1, page) = tmpColor(0) Then '下半部同
                            rv = rv & "▄" '▄的前半部
                        Else
                            rv = rv & preCode & QBC2AnsiC(MainArr(i, j + 1, page)) & "m" & "▄" & Chr(8)
                            tmpColor(0) = MainArr(i, j + 1, page)
                        End If
                    Else
                        If MainArr(i, j + 1, page) = tmpColor(0) Then
                            rv = rv & preCode & QBCToAnsiBC(MainArr(i, j, page)) & "m" & "▄" & Chr(8)
                        Else
                            rv = rv & preCode & QBC2AnsiC(MainArr(i, j + 1, page)) & ";" & QBCToAnsiBC(MainArr(i, j, page)) & "m" & "▄" & Chr(8)
                            tmpColor(0) = MainArr(i, j + 1, page)
                        End If
                        tmpColor(1) = MainArr(i, j, page)
                    End If
                    '考慮右邊
                    If MainArr(i + 1, j, page) = tmpColor(1) Then '上半部同
                        If MainArr(i + 1, j + 1, page) = tmpColor(0) Then '下半部同(在此不可能發生 已於之前排除)
                            rv = rv & "e" '▄的後半部
                            Debug.Print "發生不可能之可能"
                            
                        Else
                            rv = rv & preCode & QBC2AnsiC(MainArr(i + 1, j + 1, page)) & "m" & "e"
                            tmpColor(0) = MainArr(i + 1, j + 1, page)
                        End If
                    Else
                        If MainArr(i + 1, j + 1, page) = tmpColor(0) Then
                            rv = rv & preCode & QBCToAnsiBC(MainArr(i + 1, j, page)) & "m" & "e"
                        Else
                            rv = rv & preCode & QBC2AnsiC(MainArr(i + 1, j + 1, page)) & ";" & QBCToAnsiBC(MainArr(i + 1, j, page)) & "m" & "e"
                            tmpColor(0) = MainArr(i + 1, j + 1, page)
                        End If
                        tmpColor(1) = MainArr(i + 1, j, page)
                    End If
                End If
                
                '因考慮了兩個位元 故跳過下一個
                i = i + 1
            End If
        Next i
        '以下為修正畫面破裂程式碼
        If FixBlankEnd = 1 Then
            If tmpColor(1) <> 0 Then
                rv = rv & preCode & "m" & vbCrLf
                tmpColor(0) = 7 '前景淺灰
                tmpColor(1) = 0 '背景黑
            Else
                If deColorLost = 1 Then
                    tmpColor(0) = 7 '前景淺灰
                '    tmpColor(1) = 0 '背景黑
                    rv = rv & vbCrLf & dCLstr
                Else
                    rv = rv & vbCrLf
                End If
               
            End If
        Else
            If deColorLost = 1 Then
                tmpColor(0) = 7 '前景淺灰
                tmpColor(1) = 0 '背景黑
                rv = rv & vbCrLf & preCode & "m"
            Else
                rv = rv & vbCrLf
            End If
            
        End If
    Next j
    rv = rv & preCode & "m"
    'GetAnsiText_Area = rv

End Function
Public Function GetAnsiText_Area_old(ByVal x1 As Integer, ByVal Y1 As Integer, ByVal X2 As Integer, ByVal Y2 As Integer, ByVal page As Integer)
    Dim rv As String    '回傳暫存
    Dim tmpColor(1) As Byte '0為前景 1為背景
    Dim preCode As String
    'Dim PageLine As String '頁與頁之間的插入行(為了動畫設計)

    preCode = Chr(27) & "["
    
'    If VAFileH.filetype = 4 Then
'        PageLine = vbCrLf
'    Else
'        PageLine = vbCrLf & "---" & vbCrLf
'    End If
    
    'QBCToAnsiBC
    '預設色彩
    tmpColor(0) = 7 '前景淺灰
    tmpColor(1) = 0 '背景黑
    '進行轉換
    rv = rv & preCode & "m"
    For j = Y1 To (Y2 + 1) - 2 Step 2  '一個字元的高度為2
        For i = x1 To X2
            If MainArr(i, j, page) = MainArr(i, j + 1, page) Then
                '是否與之前的背景同色(減肥用)
                If MainArr(i, j, page) = tmpColor(1) Then
                    rv = rv & " "
                Else
                    rv = rv & preCode & QBCToAnsiBC(MainArr(i, j, page)) & "m "
                    tmpColor(1) = MainArr(i, j, page)
                End If
            Else
                '四格一起考慮
                If MainArr(i, j, page) = MainArr(i + 1, j, page) And MainArr(i, j + 1, page) = MainArr(i + 1, j + 1, page) Then
                    If MainArr(i, j, page) = tmpColor(1) Then '上半部同
                        If MainArr(i, j + 1, page) = tmpColor(0) Then '下半部同
                            rv = rv & "▄"
                        Else
                            rv = rv & preCode & QBC2AnsiC(MainArr(i, j + 1, page)) & "m" & "▄"
                            tmpColor(0) = MainArr(i, j + 1, page)
                        End If
                    Else
                        If MainArr(i, j + 1, page) = tmpColor(0) Then
                            rv = rv & preCode & QBCToAnsiBC(MainArr(i, j, page)) & "m" & "▄"
                        Else
                            rv = rv & preCode & QBC2AnsiC(MainArr(i, j + 1, page)) & ";" & QBCToAnsiBC(MainArr(i, j, page)) & "m" & "▄"
                            tmpColor(0) = MainArr(i, j + 1, page)
                        End If
                        tmpColor(1) = MainArr(i, j, page)
                    End If
                Else
                    '先考慮左邊
                    If MainArr(i, j, page) = tmpColor(1) Then '上半部同
                        If MainArr(i, j + 1, page) = tmpColor(0) Then '下半部同
                            rv = rv & "▄" '▄的前半部
                        Else
                            rv = rv & preCode & QBC2AnsiC(MainArr(i, j + 1, page)) & "m" & "▄" & Chr(8)
                            tmpColor(0) = MainArr(i, j + 1, page)
                        End If
                    Else
                        If MainArr(i, j + 1, page) = tmpColor(0) Then
                            rv = rv & preCode & QBCToAnsiBC(MainArr(i, j, page)) & "m" & "▄" & Chr(8)
                        Else
                            rv = rv & preCode & QBC2AnsiC(MainArr(i, j + 1, page)) & ";" & QBCToAnsiBC(MainArr(i, j, page)) & "m" & "▄" & Chr(8)
                            tmpColor(0) = MainArr(i, j + 1, page)
                        End If
                        tmpColor(1) = MainArr(i, j, page)
                    End If
                    '考慮右邊
                    If MainArr(i + 1, j, page) = tmpColor(1) Then '上半部同
                        If MainArr(i + 1, j + 1, page) = tmpColor(0) Then '下半部同(在此不可能發生 已於之前排除)
                            rv = rv & "e" '▄的後半部
                            Debug.Print "發生不可能之可能"
                            
                        Else
                            rv = rv & preCode & QBC2AnsiC(MainArr(i + 1, j + 1, page)) & "m" & "e"
                            tmpColor(0) = MainArr(i + 1, j + 1, page)
                        End If
                    Else
                        If MainArr(i + 1, j + 1, page) = tmpColor(0) Then
                            rv = rv & preCode & QBCToAnsiBC(MainArr(i + 1, j, page)) & "m" & "e"
                        Else
                            rv = rv & preCode & QBC2AnsiC(MainArr(i + 1, j + 1, page)) & ";" & QBCToAnsiBC(MainArr(i + 1, j, page)) & "m" & "e"
                            tmpColor(0) = MainArr(i + 1, j + 1, page)
                        End If
                        tmpColor(1) = MainArr(i + 1, j, page)
                    End If
                End If
                
                '因考慮了兩個位元 故跳過下一個
                i = i + 1
            End If
        Next i
        rv = rv & vbCrLf
    Next j
    rv = rv & preCode & "m"
    GetAnsiText_Area_old = rv

End Function
Public Function QBC2AnsiC(ByVal QBC As Integer) As String
    Dim AnsiC As String
    
    Select Case QBC Mod 8
        Case Is = 0
            AnsiC = "30"
        Case Is = 1
            AnsiC = "34"
        Case Is = 2
            AnsiC = "32"
        Case Is = 3
            AnsiC = "36"
        Case Is = 4
            AnsiC = "31"
        Case Is = 5
            AnsiC = "35"
        Case Is = 6
            AnsiC = "33"
        Case Is = 7
            AnsiC = "37"
    End Select

    QBC2AnsiC = AnsiC
End Function

'###檔案處理###
'基礎函數
'不擺在VA_File模組內是為了可以獨立抽出而不依賴VA_File模組
Private Sub VAC_SaveFile(ByVal filepath As String, ByRef bytearr() As Byte, ByRef VAhead As VAFileInfo)
On Error GoTo out
    '如果檔案已經存在就把他砍掉
    If Dir(filepath, vbHidden Or vbDirectory Or vbReadOnly Or vbSystem) <> "" Then
            Kill filepath
    End If
    OpFile = 1
    Open filepath For Binary As #OpFile
    Put #OpFile, 1, VAhead
    Put #OpFile, 11, bytearr
    Debug.Print "savefile  " & filepath
    Close #OpFile
    Exit Sub
out:
    Debug.Print "VAC_SaveFile Error Out: "; Err.Description & filepath
    
End Sub

Private Function VAC_ReadFile(ByVal filepath As String, ByRef bytearr() As Byte, ByRef VAhead As VAFileInfo) As Byte
'成功回傳1
    
    On Error GoTo out
    OpFile = 1
    Open filepath For Binary As #OpFile
    Get #OpFile, 1, VAhead
    ReDim bytearr(VAhead.ArrXUbound + 1, VAhead.ArrYUbound, 1 To VAhead.ArrZLenth) '(x-1)+1 之+1是為了讓編譯時方便(因有判斷i+1之情境)
    Get #OpFile, 11, bytearr
    Close #OpFile
    VAC_ReadFile = 1
    Exit Function
out:
    Debug.Print "VAC_ReadFile Error Out: "; Err.Description
    VAC_ReadFile = 0
End Function

Public Sub SaveFile(ByVal filepath As String)
    VAFileH.IDC = 827
    Call VAC_SaveFile(filepath, MainArr(), VAFileH)
    If cFilepath = "" Then cFilepath = filepath
End Sub

Public Sub OpenFile(ByVal filepath As String)
    PB.Parent.ScaleMode = 3
    If VAC_ReadFile(filepath, MainArr(), VAFileH) = 1 Then
        PB.Width = PPCL * VAFileH.ArrXUbound + 15     '=PPCL * (VAFileH.ArrXUbound+1) + 5   5為邊框修正量
        PB.Height = PPCL * VAFileH.ArrYUbound + 15    '=PPCL * (VAFileH.ArrYUbound+1) + 5
        cFilepath = filepath                        '設定路徑
        cPage = 1 '設定頁數預設值
        Call ReShow(0, 0, VAFileH.ArrXUbound, VAFileH.ArrYUbound)   '更新畫面
    End If
End Sub

Public Sub NewFile(ByVal Wight As Integer, ByVal Height As Integer, ByVal Z As Integer, filetype As Byte)
    Call SetSize(Wight, Height, Z)
    VAFileH.IDC = 827
    VAFileH.filetype = filetype
    cFilepath = ""
End Sub

Public Sub ReFreshColor()
    SelectObject PB.hdc, hPenColor(CurrentColor)
    SelectObject PB.hdc, hBrushColor(CurrentColor)
End Sub

'####處理多頁####
Public Sub GoToPage(ByVal pageNum As Integer)
    If pageNum <= VAFileH.ArrZLenth Then
        cPage = pageNum
        Call ReShow(0, 0, VAFileH.ArrXUbound, VAFileH.ArrYUbound)   '更新畫面
    End If
    
End Sub

Public Sub InsertPage(ByVal pageNum As Integer)
    VAFileH.ArrZLenth = VAFileH.ArrZLenth + 1
    ReDim Preserve MainArr(VAFileH.ArrXUbound + 1, VAFileH.ArrYUbound, 1 To VAFileH.ArrZLenth)

    If pageNum <> VAFileH.ArrZLenth Then
        For k = VAFileH.ArrZLenth To pageNum + 1 Step -1
            For j = 0 To VAFileH.ArrYUbound
                For i = 0 To VAFileH.ArrXUbound
                    MainArr(i, j, k) = MainArr(i, j, k - 1)
                Next i
            Next j
        Next k
        
        For j = 0 To VAFileH.ArrYUbound
            For i = 0 To VAFileH.ArrXUbound
                MainArr(i, j, pageNum) = 0
            Next i
        Next j
    End If
End Sub

'####畫#####
Public Sub Draw_Point(ByVal X As Integer, ByVal Y As Integer)
On Error GoTo out
    '裡層記憶體繪圖
    MainArr(X, Y, cPage) = CurrentColor
    '表層顯示繪圖
    Call DrawDC_Point(X, Y)
    PB.Refresh
    Exit Sub
out:
Debug.Print "Draw_Point error out"
End Sub
Public Sub Draw_Area(ByVal x1 As Integer, ByVal Y1 As Integer, ByVal X2 As Integer, ByVal Y2 As Integer)
On Error GoTo out
    'PB.AutoRedraw = False
    For j = Y1 To Y2
        For i = x1 To X2
            MainArr(i, j, cPage) = CurrentColor
            Call DrawDC_Point(i, j)
           
        Next i
    Next j

    
    'PB.AutoRedraw = True
    PB.Refresh
    Exit Sub
    
out:
Debug.Print "Draw_Point error out"
End Sub
Public Sub Erease_Point(ByVal X As Integer, ByVal Y As Integer)
On Error GoTo out
    '裡層記憶體繪圖
    MainArr(X, Y, cPage) = CurrentColor
    '表層顯示繪圖
    'Call DrawDC_Point(X, Y)
    PB.Refresh
    Exit Sub
out:
Debug.Print "Draw_Point error out"
End Sub

'#####換色#####


Public Sub ExChColor_Point(ByVal X As Integer, ByVal Y As Integer, ByVal fromColor As Byte, ByVal toColor As Byte)
On Error GoTo out
    If MainArr(X, Y, cPage) = fromColor Then
        MainArr(X, Y, cPage) = toColor
        Call ReShow_Point(X, Y)
        PB.Refresh
        'Debug.Print "x,y=" & X & "," & Y & " fromColor/toColor :" & fromColor & "/" & toColor
    End If
    Exit Sub
out:
    Debug.Print "ExChColor_Point Error Out" & Err.Description
End Sub

Public Sub ExChColor_Area(ByVal x1 As Integer, ByVal Y1 As Integer, ByVal X2 As Integer, ByVal Y2 As Integer, ByVal fromColor As Byte, ByVal toColor As Byte)
    If x1 < 0 Then x1 = 0
    If Y1 < 0 Then Y1 = 0
    If X2 > VAFileH.ArrXUbound Then X2 = VAFileH.ArrXUbound
    If Y2 > VAFileH.ArrYUbound Then Y2 = VAFileH.ArrYUbound
    'PB.AutoRedraw = False
    For j = Y1 To Y2
        For i = x1 To X2
            If MainArr(i, j, cPage) = fromColor Then
                MainArr(i, j, cPage) = toColor
                Call ReShow_Point(i, j)
            End If
        Next i
        'Debug.Print "line :" & j
    Next j
    PB.Refresh
    'PB.AutoRedraw = True
End Sub

'####清除頁面####


Public Sub Erease_All()
    PB.AutoRedraw = False
    For j = 0 To VAFileH.ArrYUbound
        For i = 0 To VAFileH.ArrXUbound
            MainArr(i, j, cPage) = 0
        Next i
        'Debug.Print "line :" & j
    Next j
    PB.AutoRedraw = True
    PB.Cls
    Call ReSetDC
    If IP_Vars(10) = 1 Then Call ReShow_BG(0, 0, VAFileH.ArrXUbound, VAFileH.ArrYUbound)
    
End Sub

Public Sub Erease_Area(ByVal x1 As Integer, ByVal Y1 As Integer, ByVal X2 As Integer, ByVal Y2 As Integer)
    Dim tmpbyte As Byte
    tmpbyte = CurrentColor
    If x1 < 0 Then x1 = 0
    If Y1 < 0 Then Y1 = 0
    Call SetColor(0)
    If IP_Vars(10) = 1 Then
        Call ReShow_BG(x1, Y1, X2, Y2)
        For j = Y1 To Y2
            For i = x1 To X2
                MainArr(i, j, cPage) = 0
            Next i
        Next j
    Else
        Call Draw_Area(x1, Y1, X2, Y2)
    End If
    
    'PB.AutoRedraw = True
    PB.Refresh
    Call SetColor(tmpbyte)
End Sub

'#####特殊效果####

Public Sub Flip_V(ByVal x1 As Integer, ByVal Y1 As Integer, ByVal X2 As Integer, ByVal Y2 As Integer)
    '垂直翻轉

    Dim tmpYlen As Integer
    Dim tmpbyte As Byte
    If x1 < 0 Then x1 = 0
    If Y1 < 0 Then Y1 = 0
    If X2 > VAFileH.ArrXUbound Then X2 = VAFileH.ArrXUbound
    If Y2 > VAFileH.ArrYUbound Then Y2 = VAFileH.ArrYUbound
    
    tmpYlen = (Y2 - Y1 + 1) \ 2 - 1
    If tmpYlen < 0 Then Exit Sub
    
    For j = 0 To tmpYlen
        For i = x1 To X2
            tmpbyte = MainArr(i, Y1 + j, cPage)
            MainArr(i, Y1 + j, cPage) = MainArr(i, Y2 - j, cPage)
            MainArr(i, Y2 - j, cPage) = tmpbyte
        Next i
    Next j
    Call ReShow(x1, Y1, X2, Y2)
End Sub

Public Sub Flip_H(ByVal x1 As Integer, ByVal Y1 As Integer, ByVal X2 As Integer, ByVal Y2 As Integer)
    '水平翻轉
    Dim tmpXlen As Integer
    If x1 < 0 Then x1 = 0
    If Y1 < 0 Then Y1 = 0
    If X2 > VAFileH.ArrXUbound Then X2 = VAFileH.ArrXUbound
    If Y2 > VAFileH.ArrYUbound Then Y2 = VAFileH.ArrYUbound
    tmpXlen = (X2 - x1) \ 2
    If tmpXlen < 0 Then Exit Sub
    For j = Y1 To Y2
        For i = 0 To tmpXlen
            tmpbyte = MainArr(x1 + i, j, cPage)
            MainArr(x1 + i, j, cPage) = MainArr(X2 - i, j, cPage)
            MainArr(X2 - i, j, cPage) = tmpbyte
        Next i
    Next j
    Call ReShow(x1, Y1, X2, Y2)
End Sub

Public Sub FlipColor(ByVal x1 As Integer, ByVal Y1 As Integer, ByVal X2 As Integer, ByVal Y2 As Integer)
    If x1 < 0 Then x1 = 0
    If Y1 < 0 Then Y1 = 0
    If X2 > VAFileH.ArrXUbound Then X2 = VAFileH.ArrXUbound
    If Y2 > VAFileH.ArrYUbound Then Y2 = VAFileH.ArrYUbound
    For j = Y1 To Y2
        For i = x1 To X2
            MainArr(i, j, cPage) = 7 - MainArr(i, j, cPage)
        Next i
    Next j
    Call ReShow(x1, Y1, X2, Y2)
End Sub


Public Sub Rotate_Right(ByVal x1 As Integer, ByVal Y1 As Integer, ByVal X2 As Integer, ByVal Y2 As Integer)
    '像右旋轉
    On Error Resume Next
    Dim tmparr() As Byte, dX As Integer, dY As Integer
    dX = X2 - x1
    dY = Y2 - Y1
    If (x1 + dY) > VAFileH.ArrXUbound Then dY = (VAFileH.ArrXUbound - x1)
    If (Y1 + dX) > VAFileH.ArrYUbound Then dX = (VAFileH.ArrYUbound - Y1)
    ReDim tmparr(dY, dX)
    For j = 0 To dY
        For i = 0 To dX
            tmparr(dY - j, i) = MainArr(x1 + i, Y1 + j, cPage)
        Next i
    Next j
    Call Erease_Area(x1, Y1, X2, Y2)
    For j = 0 To dX
        For i = 0 To dY
            MainArr(x1 + i, Y1 + j, cPage) = tmparr(i, j)
        Next i
    Next j
    If dX > dY Then
        Call ReShow(x1, Y1, x1 + dX, Y1 + dX)
    Else
        Call ReShow(x1, Y1, x1 + dY, Y1 + dY)
    End If
    
End Sub

'#####剪貼簿功能(CPB)#######

Public Sub CPB_Copy(ByVal x1 As Integer, ByVal Y1 As Integer, ByVal X2 As Integer, ByVal Y2 As Integer)
    Dim dX As Integer, dY As Integer
    If x1 < 0 Then x1 = 0
    If Y1 < 0 Then Y1 = 0
    If X2 > VAFileH.ArrXUbound Then
        dX = VAFileH.ArrXUbound - x1
    Else
        dX = X2 - x1
    End If
    If Y2 > VAFileH.ArrYUbound Then
        dX = VAFileH.ArrYUbound - Y1
    Else
        dY = Y2 - Y1
    End If
    
    CPB_uX = dX
    CPB_uY = dY
    ReDim CPBArr(dX, dY)
    For j = 0 To dY
        For i = 0 To dX
            CPBArr(i, j) = MainArr(x1 + i, Y1 + j, cPage)
            
        Next i
    Next j
    
End Sub
Public Sub CPB_Cut(ByVal x1 As Integer, ByVal Y1 As Integer, ByVal X2 As Integer, ByVal Y2 As Integer)
    Call CPB_Copy(x1, Y1, X2, Y2)
    Call Erease_Area(x1, Y1, X2, Y2)
    
End Sub
Public Sub CPB_Past(ByVal X As Integer, ByVal Y As Integer)
On Error Resume Next
    Dim tmpuX As Integer, tmpuY As Integer
    If X + CPB_uX > VAFileH.ArrXUbound Then
        tmpuX = VAFileH.ArrXUbound - X
    Else
        tmpuX = CPB_uX
    End If
    If Y + CPB_uY > VAFileH.ArrYUbound Then
        tmpuY = VAFileH.ArrYUbound - Y
    Else
        tmpuY = CPB_uY
    End If
    For j = 0 To tmpuY
        For i = 0 To tmpuX
             MainArr(X + i, Y + j, cPage) = CPBArr(i, j)
        Next i
    Next j
    Call ReShow(X, Y, X + tmpuX, Y + tmpuY)
End Sub
Public Sub CPB_Past_DeBackGround(ByVal X As Integer, ByVal Y As Integer)
On Error Resume Next
    Dim tmpuX As Integer, tmpuY As Integer
    If X + CPB_uX > VAFileH.ArrXUbound Then
        tmpuX = VAFileH.ArrXUbound - X
    Else
        tmpuX = CPB_uX
    End If
    If Y + CPB_uY > VAFileH.ArrYUbound Then
        tmpuY = VAFileH.ArrYUbound - Y
    Else
        tmpuY = CPB_uY
    End If
    For j = 0 To tmpuY
        For i = 0 To tmpuX
             If CPBArr(i, j) <> 0 Then
                  MainArr(X + i, Y + j, cPage) = CPBArr(i, j)
             End If
        Next i
    Next j
    Call ReShow(X, Y, X + tmpuX, Y + tmpuY)
End Sub
Public Sub CPB_GetWH(ByRef W As Integer, ByRef H As Integer)
    
End Sub


'####處理外部圖片####

Public Sub LoadIP(ByVal filepath As String)
On Error GoTo out
    Set IP_Out = LoadPicture(filepath)
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
    IP_outhDC = CreateCompatibleDC(PB.hdc) '建立符合主畫布的dc
    'Debug.Print "建立dc"
    SelectObject IP_outhDC, IP_Out.handle
    'Debug.Print "dc選擇bitmap:" & IP_outhDC & "<-SEL->" & IP_Out.Handle
End Sub

Public Sub ShowIPout2PB()
    'Dim tmpL(3) As Long '0
    IP_PB.Cls
    
    'Call DrawBitmap2DC(thdc, IP_Out.Handle, 0, 0)
    'StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
    StretchBlt IP_PB.hdc, IP_Vars(0), IP_Vars(1), IP_Vars(2), IP_Vars(3), IP_outhDC, IP_Vars(6), IP_Vars(7), IP_Vars(8), IP_Vars(9), vbSrcCopy
    'Debug.Print "StretchBlt " & IP_PB.hdc & "," & IP_Vars(0) & "," & IP_Vars(1) & "," & IP_Vars(2) & "," & IP_Vars(3) & "," & IP_Vars(4) & "," & "0, 0," & IP_Vars(4) & "," & IP_Vars(5) & ",vbSrcCopy"
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

Public Sub ReShow_BG(ByVal x1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long)
    '改良版的顯示背景
    'Dim tmphDC As Long
    x1 = x1 * PPCL
    Y1 = Y1 * PPCL
    X2 = X2 * PPCL
    Y2 = Y2 * PPCL
    'tmphDC = CreateCompatibleDC(hdc)  '建立相容的於目標dc的dc
    'SelectObject tmphDC, IP_Out.Handle
    BitBlt PB.hdc, x1, Y1, X2 - x1 + PPCL, Y2 - Y1 + PPCL, IP_PB.hdc, x1, Y1, vbSrcCopy
    'Debug.Print "BitBlt:"; &HDC & "," & x1 & "," & Y1 & "," & X2 - x1 & "," & Y2 - Y1 & "," & tmphDC & "," & x1 & "," & Y1 & ","
    'DeleteDC tmphDC
    
End Sub

Public Function GetIP_Vars(ByVal Index As Byte) As Integer
    '紀錄圖片縮放及位置 0:X座標 1:Y座標 2:寬 3:高 ......詳見宣告說明
    GetIP_Vars = IP_Vars(Index)
End Function

Public Sub GetBitmapWH(ByVal hbitmap As Long)
    Dim bm As BITMAP
    GetObject hbitmap, Len(bm), bm
    IP_Vars(4) = bm.bmWidth
    IP_Vars(5) = bm.bmHeight
    Debug.Print "BITMAP INFO " & bm.bmWidthBytes & ";" & bm.bmPlanes & ";" & bm.bmBitsPixel
End Sub

Public Sub Load_IP_PB(ByVal PB As PictureBox)
    Set IP_PB = PB
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
    If (VAFileH.ArrXUbound + 1) * IP_Vars(9) < (VAFileH.ArrYUbound + 1) * IP_Vars(8) Then
        '以寬為準
        IP_Vars(2) = (VAFileH.ArrXUbound + 1) * PPCL
        Debug.Print (IP_Vars(2) * IP_Vars(9)) / IP_Vars(8)
        IP_Vars(3) = Fix((IP_Vars(2) * IP_Vars(9)) / IP_Vars(8))

    Else
        '以高為準
        IP_Vars(3) = (VAFileH.ArrYUbound + 1) * PPCL
        IP_Vars(2) = Fix((IP_Vars(3) * IP_Vars(8)) / IP_Vars(9))

    End If
    Call ShowIPout2PB
End Sub
Public Sub IPED_SizePixelCube()
    '將像素放大到跟方格一樣大
    'IP_Vars(0) = IP_Vars(0) * PPCL
    'IP_Vars(1) = IP_Vars(1) * PPCL
    IP_Vars(2) = IP_Vars(8) * PPCL
    IP_Vars(3) = IP_Vars(9) * PPCL

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
Public Sub IPED_TransForm(ByVal x1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long)
    '變形
    IP_Vars(0) = x1
    IP_Vars(1) = Y1
    IP_Vars(2) = X2 - x1 + 1
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
Public Sub Pic2VAC_win(Optional deBlack As Byte)
    'win內建轉換法來自vb心得筆記
    '內建大小縮放
    '內建顏色轉換
    Dim bm As BITMAP, SizeOfArray As Long, Fp As Long
    'Dim bf As BitmapFileHeader
    Dim bi As BitMapInfo16, buffer() As Byte
    Dim tmphDC As Long, hDIB As Long, OldObj As Long
    Dim i As Long, j As Long, r As Integer, g As Integer, B As Integer
    Dim tmpL(1) As Long
    Dim tmpbyte As Byte
    Call GetObject(IP_PB.Image, Len(bm), bm)
    Debug.Print "IP_PB.Picture:" & IP_PB.Picture
    SizeOfArray = (((bm.bmWidth / 2 + 3) \ 4) * 4) * bm.bmHeight
    
    'With bf
    '    .bfType = "BM"
    '    .bfSize = Len(bf) + Len(bi) + SizeOfArray
    '    .bfReserved1 = 0
    '    .bfReserved2 = 0
    '    .bfOffBits = Len(bf) + Len(bi)
    'End With
    With bi
        With .bmiHeader
        .biSize = Len(bi.bmiHeader)
        .biWidth = bm.bmWidth
        .biHeight = bm.bmHeight
        .biPlanes = 1
        .biBitCount = 4
        .biCompression = 0
        .biSizeImage = SizeOfArray
        Debug.Print "SizeOfArray:" & SizeOfArray
        End With
        For i = 0 To 15
            .bmiColors(i) = QBColor(i)
            'Debug.Print "QBColor(" & i & ")=" & QBColor(i)
        Next i
    End With
    ReDim buffer(bi.bmiHeader.biSizeImage - 1) As Byte
    
    tmphDC = CreateCompatibleDC(0&)
    hDIB = CreateDIBSection16(tmphDC, bi, DIB_RGB_COLORS, i, 0&, 0&)
    OldObj = SelectObject(tmphDC, hDIB)
    'Call BitBlt(tmphDC, 0&, 0&, bm.bmWidth, bm.bmHeight, pic.hDC, 0&, 0&, vbSrcCopy)
    tmpL(0) = (VAFileH.ArrXUbound + 1)
    tmpL(1) = (VAFileH.ArrYUbound + 1)
    
    Call StretchBlt(tmphDC, 0, 0, tmpL(0), tmpL(1), IP_PB.hdc, 0, 0, tmpL(0) * PPCL, tmpL(1) * PPCL, vbSrcCopy)
    'Call GetDIBits16(hDC, hDIB, 0, (bm.bmHeight), buffer(0), bi, 0)
    If deBlack = 1 Then '是否去背
        For j = 0 To VAFileH.ArrYUbound
            For i = 0 To VAFileH.ArrXUbound
                tmpbyte = LC2QBC(GetPixel(tmphDC, i, j))
                If tmpbyte <> 0 Then MainArr(i, j, cPage) = tmpbyte
            Next i
        Next j
    Else
        For j = 0 To VAFileH.ArrYUbound
            For i = 0 To VAFileH.ArrXUbound
                MainArr(i, j, cPage) = LC2QBC(GetPixel(tmphDC, i, j))
            Next i
        Next j
    End If
    Call ReShow(0, 0, VAFileH.ArrXUbound, VAFileH.ArrYUbound)
    'Call StretchBlt(hdc, 0, 0, bm.bmWidth, bm.bmHeight, tmphDC, 0, 0, tmpL(0), tmpL(1), vbSrcCopy)
    'PB.Refresh
    SelectObject tmphDC, OldObj
    DeleteDC tmphDC
    DeleteObject hDIB
End Sub

Public Sub Pic2VAC_V1(Optional deBlack As Byte)
    '自設的轉換函式
    Dim tmphDC As Long, tmphBitmap As Long
    Dim tmpL(1) As Long
    tmphDC = CreateCompatibleDC(IP_PB.hdc)
    tmphBitmap = CreateCompatibleBitmap(IP_PB.hdc, PPCL * (VAFileH.ArrXUbound + 1), PPCL * (VAFileH.ArrYUbound + 1))
    SelectObject tmphDC, tmphBitmap
    tmpL(0) = (VAFileH.ArrXUbound + 1)
    tmpL(1) = (VAFileH.ArrYUbound + 1)
    Call StretchBlt(tmphDC, 0, 0, tmpL(0), tmpL(1), IP_PB.hdc, 0, 0, tmpL(0) * PPCL, tmpL(1) * PPCL, vbSrcCopy)
    Dim tmpbyte As Byte
    If deBlack = 1 Then
        For j = 0 To VAFileH.ArrYUbound
            For i = 0 To VAFileH.ArrXUbound
                tmpbyte = ColorAnalize_v2(GetPixel(tmphDC, i, j))
                If tmpbyte <> 0 Then MainArr(i, j, cPage) = tmpbyte
            Next i
        Next j
    Else
        For j = 0 To VAFileH.ArrYUbound
            For i = 0 To VAFileH.ArrXUbound
                MainArr(i, j, cPage) = ColorAnalize_v2(GetPixel(tmphDC, i, j))
            Next i
        Next j
    End If
    Call ReShow(0, 0, VAFileH.ArrXUbound, VAFileH.ArrYUbound)
    DeleteDC tmphDC
    DeleteObject tmphBitmap
End Sub

Public Function LC2QBC(LC As Long) As Byte
    '將以long為紀錄的顏色資訊 轉換為qbcolor的輸入值(0~15)
    Dim rv As Byte
    Select Case LC
        Case QBLC(0), QBLC(8)
            rv = 0
        Case QBLC(1), QBLC(9)
            rv = 1
        Case QBLC(2), QBLC(10)
            rv = 2
        Case QBLC(3), QBLC(11)
            rv = 3
        Case QBLC(4), QBLC(12)
            rv = 4
        Case QBLC(5), QBLC(13)
            rv = 5
        Case QBLC(6), QBLC(14)
            rv = 6
        Case QBLC(7), QBLC(15)
            rv = 7
        Case Else
            Debug.Print "LC2QBC發生例外 : " & LC
        
    End Select
    
    LC2QBC = rv
End Function

Private Function ColorAnalize_v1(ByVal colorLong As Long) As Byte
    '自製顏色分析函式1
    
    '色彩閥值
    Dim tv_R As Byte, tv_G As Byte, tv_B As Byte
    tv_R = 100
    tv_G = 100
    tv_B = 100
    Dim tmp_R As Byte, tmp_G As Byte, tmp_B As Byte
    'BGR
    tmp_B = colorLong \ (256 ^ 2)
    tmp_G = (colorLong Mod (256 ^ 2)) \ 256
    tmp_R = colorLong Mod 256
    Dim qbcc As Byte
    'Debug.Print "R:G:B =" & tmp_R & ":" & tmp_G & ":" & tmp_B
    'qbcc = (tmp_R \ tv_R) * 2 ^ 2 + (tmp_G \ tv_G) * 2 ^ 1 + (tmp_B \ tv_B) * 2 ^ 0
    qbcc = CA_BS10(tmp_R, tv_R) * 2 ^ 2 + CA_BS10(tmp_G, tv_G) * 2 ^ 1 + CA_BS10(tmp_B, tv_B) * 2 ^ 0
    'Debug.Print qbcc
    ColorAnalize_v1 = qbcc
    
End Function

Private Function CA_BS10(ByVal a As Byte, ByVal B As Byte) As Byte
    'ColorAnalize Big Small 1 0
    'a大於等於b則1 a小於b則0
    If a >= B Then
        CA_BS10 = 1
    Else
        CA_BS10 = 0
    End If
End Function
Private Function ColorAnalize_v2(ByVal colorLong As Long) As Byte
    '自製顏色分析函式1
    

    'Dim tmp_R As Integer, tmp_G As Integer, tmp_B As Integer
    'BGR
    CA_Vars(2) = colorLong \ (256 ^ 2)
    CA_Vars(1) = (colorLong Mod (256 ^ 2)) \ 256
    CA_Vars(0) = colorLong Mod 256
    '色彩閥值
    'Dim tv_R As Byte, tv_G As Byte, tv_B As Byte
    CA_Vars(3) = CA_v2_BC((CA_Vars(2) + CA_Vars(1) + CA_Vars(0)) \ 3)
    CA_Vars(4) = CA_v2_BC((CA_Vars(2) + CA_Vars(1) + CA_Vars(0)) \ 3)
    CA_Vars(5) = CA_v2_BC((CA_Vars(2) + CA_Vars(1) + CA_Vars(0)) \ 3)
    'Dim qbcc As Byte
    'Debug.Print "R:G:B =" & tmp_R & ":" & tmp_G & ":" & tmp_B
    'qbcc = (tmp_R \ tv_R) * 2 ^ 2 + (tmp_G \ tv_G) * 2 ^ 1 + (tmp_B \ tv_B) * 2 ^ 0
    ColorAnalize_v2 = CA_BS10(CA_Vars(0), CA_Vars(3)) * 2 ^ 2 + CA_BS10(CA_Vars(1), CA_Vars(4)) * 2 ^ 1 + CA_BS10(CA_Vars(2), CA_Vars(5)) * 2 ^ 0
    'Debug.Print qbcc

    
End Function


Private Function CA_v2_BC(ByVal tv As Byte) As Byte
    'ColorAnalize_v2 bright Control 亮度控制
    '144~112
    If tv > CA_Vars(7) Then tv = CA_Vars(7)
    If tv < CA_Vars(6) Then tv = CA_Vars(6)
    CA_v2_BC = tv
End Function
Public Sub CA_v2_BC_DEFUALT()
    'ColorAnalize_v2 bright Control 亮度控制
    '144~112
    CA_Vars(7) = 144
    CA_Vars(6) = 112

End Sub

Public Sub CA_v2_SetCaVars(ByVal Index As Byte, ByVal Value As Byte)
    CA_Vars(Index) = Value
End Sub
Public Function CA_v2_GetCaVars(ByVal Index As Byte) As Byte
    CA_v2_GetCaVars = CA_Vars(Index)
End Function

Public Function GetPicBright(ByRef mean As Single, ByRef sd As Single) As Integer
    '取得整張圖片的亮度
    Dim tmphDC As Long, tmphBitmap As Long
    Dim tmpL(3) As Long
    
    
    tmphDC = CreateCompatibleDC(IP_PB.hdc)
    tmphBitmap = CreateCompatibleBitmap(IP_PB.hdc, PPCL * (VAFileH.ArrXUbound + 1), PPCL * (VAFileH.ArrYUbound + 1))
    SelectObject tmphDC, tmphBitmap
    tmpL(0) = (VAFileH.ArrXUbound + 1)
    tmpL(1) = (VAFileH.ArrYUbound + 1)
    Call StretchBlt(tmphDC, 0, 0, tmpL(0), tmpL(1), IP_PB.hdc, 0, 0, tmpL(0) * PPCL, tmpL(1) * PPCL, vbSrcCopy)
    Dim tmpbyte As Byte
    tmpL(2) = 0
    tmpL(1) = 0
    tmpL(3) = 0
    For j = 0 To VAFileH.ArrYUbound
        For i = 0 To VAFileH.ArrXUbound
            tmpL(0) = GetPixel(tmphDC, i, j)
            If tmpL(0) <> 0 Then
                tmpL(0) = (tmpL(0) \ (256 ^ 2) + (tmpL(0) Mod (256 ^ 2)) \ 256 + tmpL(0) Mod 256) \ 3
                tmpL(3) = tmpL(3) + tmpL(0) ^ 2
                tmpL(1) = tmpL(1) + tmpL(0)
                tmpL(2) = tmpL(2) + 1
                Debug.Print tmpL(0)

            End If
            'MainArr(i, j, cPage) = ColorAnalize_v2(GetPixel(tmphDC, i, j))
        Next i
    Next j
    If tmpL(2) <> 0 Then
        mean = tmpL(1) / tmpL(2)
        sd = ((tmpL(3) - tmpL(2) * mean ^ 2) / (tmpL(2) - 1)) ^ 0.5
        'Debug.Print "sd = ((" & tmpL(3) & " - " & tmpL(2) & " * " & mean ^ 2 & ") / (" & tmpL(2) & " - 1)) ^ 0.5"
        GetPicBright = mean
    Else
        
    End If
    
    DeleteDC tmphDC
    DeleteObject tmphBitmap
End Function

'是否使用背景
Public Property Get HaveBG() As Variant
    HaveBG = IP_Vars(10)
End Property

Public Property Let HaveBG(ByVal vNewValue As Variant)
    IP_Vars(10) = vNewValue
End Property

