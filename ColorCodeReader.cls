VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ColorCodeReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private tmpCCStr As String 'Temp Color Code String
Private tmpCLA() As ColorLayer
Private tmpVArr() As Variant
Private Ba_CLA() As ColorLayer


Public Property Get SetCCStr() As Variant
    SetCCStr = tmpCCStr
End Property

Public Property Let SetCCStr(ByVal vNewValue As Variant)
    tmpCCStr = vNewValue
    
End Property

Public Sub AnalyzeCC()
On Error GoTo out
'主陣列tmpVArr(行索引)(區段索引,0 or 1) 0存色碼陣列 1存字串
'宣告變數
    Dim tmpStrLA() As String    '儲存每行
    Dim tmpstrA() As String     '儲存各個命令區段
    Dim tmpStrA1() As String    '儲存0命令1字串
    Dim tmpInt As Integer       '純粹
    Dim IntCounter As Integer
    Dim tmpIntCounter As Integer
    Dim tmpbyte As Byte
    Dim tmpStr As String * 1
'第一階段建立STR主陣列
    tmpStrLA = Split(tmpCCStr, vbCrLf) '將行分開
    tmpInt = UBound(tmpStrLA)   'tmpint為行數-1
    ReDim tmpVArr(tmpInt)    '建立主陣列
    For i = 0 To tmpInt
        If tmpStrLA(i) <> "" Then
            tmpstrA = Split(tmpStrLA(i), "[")  '指令區段分離
            ReDim tmpStrA1(UBound(tmpstrA), 1)  '第二維的 0存色碼陣列 1存字串
            tmpIntCounter = 0  '將計算字數的counter歸零
            For j = 0 To UBound(tmpstrA)
                tmpInt = InStr(tmpstrA(j), "m")
                If tmpInt <> 0 Then
                    tmpStrA1(j, 0) = Replace(Left(tmpstrA(j), tmpInt - 1), " ", "") '題取出色彩碼 並去除所有空白
                End If
                tmpStrA1(j, 1) = Right(tmpstrA(j), Len(tmpstrA(j)) - tmpInt)    '題取出ANSI字串部分
                tmpIntCounter = tmpIntCounter + Tlen(tmpStrA1(j, 1))    '將字數累計到COUNTER中
                tmpVArr(i) = tmpStrA1   '將此行之陣列放入主陣列
                
            Next j
        Else
            ReDim tmpStrA1(0, 1)
            tmpVArr(i) = tmpStrA1
        End If
        If tmpIntCounter > IntCounter Then IntCounter = tmpIntCounter   '如果此行字數的比較多就設定寬為此行行數
    Next i
    If IntCounter = 0 Then
        Debug.Print "error 匯入彩色碼貼上的寬度是0  "
        Exit Sub
    End If
'第二階段將STR陣列製成COLOR LAYER ARRAY
    Dim tmpStrA3() As String '分析色碼的暫存容器
    Dim tmpCountPerLine As Integer  '行中字數指標(tlen)
    Dim tmpCL As ColorLayer
    Dim tmpCLbyte(2) As Byte
    If IntCounter = 0 Then Debug.Print "錯誤: 寬度不可為零"
    ReDim tmpCLA(IntCounter - 1, UBound(tmpVArr)) '建立COLOR LAYER陣列
    Debug.Print "最寬寬度: " & IntCounter
    tmpCL.BColor = 0
    tmpCL.Color = 7
    For i = 0 To UBound(tmpVArr)
        tmpCountPerLine = 1 '將指標歸零
        For j = 0 To UBound(tmpVArr(i), 1)
            'Debug.Print "Line:" & i; " ;Section: " & j & ";Len=" & Tlen(tmpVArr(i)(j, 1)) & "Pos:" & tmpCountPerLine
            If tmpVArr(i)(j, 0) = "" And j <> 0 Then
                '色彩全部還原為預設值
                tmpCL.BColor = 0
                tmpCL.Color = 7
            Else
                tmpStrA3 = Split(tmpVArr(i)(j, 0), ";") '將各色碼分離
                If UBound(tmpStrA3) = 0 Then
                '當色碼只有一個時
                    tmpbyte = Val(tmpStrA3(0))
                    Select Case tmpbyte
                        Case 1
                            If tmpCL.Color \ 8 = 0 Then tmpCL.Color = tmpCL.Color + 8
                        Case Is >= 40
                            tmpCL.BColor = GetBC(tmpbyte)
                        Case Is >= 30
                            tmpCL.Color = GetFC(tmpbyte) + (tmpCL.Color \ 8) * 8
                    End Select
                Else
                    'Debug.Print "多色碼:" & UBound(tmpStrA3)
                    tmpCLbyte(0) = 0
                    tmpCLbyte(1) = 255
                    tmpCLbyte(2) = 255
                    For k = 0 To UBound(tmpStrA3)
                    '先處理色彩
                        tmpbyte = Val(tmpStrA3(k))
                        '設定暫存之暫存預設 255的數字只是為了跟0作有沒有變更的分辯用

                        Select Case tmpbyte
                            '還原色彩暫存為預設
                            Case 0
                                tmpCL.BColor = 0
                                tmpCL.Color = 7
                                'Debug.Print "暫存色彩還原預設"
                            '運算色彩暫存之暫存 用於比較
                            Case Is = 1
                                tmpCLbyte(0) = 8
                            Case Is >= 40
                                tmpCLbyte(2) = GetBC(tmpbyte)
                            Case Is >= 30
                                tmpCLbyte(1) = GetFC(tmpbyte)
                        End Select
                    Next k
                    '從(色彩暫存之暫存)寫入(色彩暫存)
                    If tmpCLbyte(1) <> 255 Then
                        tmpCL.Color = tmpCLbyte(1) + (tmpCL.Color \ 8) * 8
                    End If
                    If tmpCLbyte(0) <> 0 And tmpCL.Color \ 8 = 0 Then
                        tmpCL.Color = tmpCL.Color + tmpCLbyte(0)
                        'Debug.Print "亮色"
                    End If
                    If tmpCLbyte(2) <> 255 Then
                        tmpCL.BColor = tmpCLbyte(2)
                    End If
                    
                End If

            End If
                    '再寫入文字
                    If tmpVArr(i)(j, 1) <> "" Then
                        For L = 1 To Len(tmpVArr(i)(j, 1))
                            tmpStr = Mid(tmpVArr(i)(j, 1), L, 1)
                            'Debug.Print "Line:" & i; " ;section: " & j & ";字數:" & tmpCountPerLine - 1
                            tmpCLA(tmpCountPerLine - 1, i).Ansi = Asc(tmpStr)
                            tmpCLA(tmpCountPerLine - 1, i).Color = tmpCL.Color
                            tmpCLA(tmpCountPerLine - 1, i).BColor = tmpCL.BColor
                            tmpbyte = Tlen(tmpStr)
                            If tmpbyte = 2 Then
                                tmpCLA(tmpCountPerLine, i).Ansi = -1
                                tmpCLA(tmpCountPerLine, i).Color = tmpCL.Color
                                tmpCLA(tmpCountPerLine, i).BColor = tmpCL.BColor
                            End If
                            tmpCountPerLine = tmpCountPerLine + tmpbyte
                            '讓CPU喘息一下
                            DoEvents
                        Next L
                    End If
        Next j
    Next i
Exit Sub
out:
    Debug.Print "AnalyzeCC Error Out : " & Err.Description
End Sub



Public Function GetFC(ByVal inFCC As Byte) As Byte
    'FCC (FORE COLOR CODE)
    '定義1;37 => 137
    Dim tmpbyte  As Byte
    
    Select Case inFCC Mod 100
        Case Is = 30
            tmpbyte = 0
        Case Is = 31
            tmpbyte = 4
        Case Is = 32
            tmpbyte = 2
        Case Is = 33
            tmpbyte = 6
        Case Is = 34
            tmpbyte = 1
        Case Is = 35
            tmpbyte = 5
        Case Is = 36
            tmpbyte = 3
        Case Is = 37
            tmpbyte = 7
        Case Else
            Debug.Print "SetFC 無法辨識來源"
    End Select
    '處理亮色
    If inFCC > 100 Then
        tmpbyte = tmpbyte + 8
    End If
    '回傳
    GetFC = tmpbyte
    
End Function
Public Function GetBC(ByVal inBCC As Byte) As Byte
    'BCC (BACK COLOR CODE)

    Dim tmpbyte  As Byte
    
    Select Case inBCC
        Case Is = 40
            tmpbyte = 0
        Case Is = 41
            tmpbyte = 4
        Case Is = 42
            tmpbyte = 2
        Case Is = 43
            tmpbyte = 6
        Case Is = 44
            tmpbyte = 1
        Case Is = 45
            tmpbyte = 5
        Case Is = 46
            tmpbyte = 3
        Case Is = 47
            tmpbyte = 7
        Case Else
            Debug.Print "GETBC 無法辨識來源"
    End Select

    '回傳
    GetBC = tmpbyte
    
End Function

Public Sub Post2Arrf(ByVal X As Integer, ByVal Y As Integer, Optional ByVal deBack As Byte)
    '辨別是否要去背
    If deBack = 1 Then
        Call CLArrayPaste_C(tmpCLA, Arrf, X, Y, OFP.CurrentPage)
    Else
        Call ObjLibPo(tmpCLA, Arrf, X, Y)
    End If
End Sub

Public Function AnalyzeCC_ByteArray(ByVal X As Integer, ByVal Y As Integer, ByVal maxX As Integer, ByVal maxY As Integer, Optional ByVal toClpB As Byte)
    '讀取剪貼簿的陣列進行分析
    Dim CCMode As Byte  '0 一般
    Dim tmpStr As String '字串暫存器
    Dim tmpInt As Integer
    Dim i As Long
    Dim Pointer_X As Integer, Pointer_Y As Integer, cutX As Integer, cutY As Integer
    Dim tmpbiByte(1) As Byte, biByteFlag As Byte, preFC_Byte As Byte, preBC_Byte As Byte, FC_Byte As Byte, BC_Byte As Byte, rawFClight_Byte As Byte, rawFC_Byte As Byte, rawBC_Byte As Byte
    Dim colorResetFlag As Byte
    Dim tmpint2 As Integer
    ReDim Ba_CLA(maxX, maxY)
    FC_Byte = 7
    BC_Byte = 0
    For i = 0 To UBound(ByteArray)
        If CCMode = 0 Then
            If ByteArray(i) = 0 Then
                Exit For
            ElseIf ByteArray(i) = 27 Then   '彩色碼認證
                If ByteArray(i + 1) = 91 Then   '確定彩色碼辨認字元後為"["就進入模式1
                    i = i + 1
                    CCMode = 1
                End If
            ElseIf ByteArray(i) = 10 Then   '\n 換行
                If Pointer_Y < maxY Then
                    Pointer_X = 0
                    Pointer_Y = Pointer_Y + 1
                End If
                biByteFlag = 0
            ElseIf biByteFlag = 1 Then
                If Pointer_X < maxX Then
                    tmpbiByte(0) = ByteArray(i)
                    Call CopyMemory(tmpInt, tmpbiByte(0), 2)
                    Ba_CLA(Pointer_X, Pointer_Y).Ansi = tmpInt
                    Ba_CLA(Pointer_X, Pointer_Y).Color = preFC_Byte
                    Ba_CLA(Pointer_X, Pointer_Y).BColor = preBC_Byte
                    Ba_CLA(Pointer_X + 1, Pointer_Y).Ansi = -1
                    Ba_CLA(Pointer_X + 1, Pointer_Y).Color = FC_Byte
                    Ba_CLA(Pointer_X + 1, Pointer_Y).BColor = BC_Byte
                    If cutX < Pointer_X + 1 Then cutX = Pointer_X + 1 '取最大的x
                    Pointer_X = Pointer_X + 2
                End If
                biByteFlag = 0
            ElseIf (ByteArray(i) >= 32 And ByteArray(i) <= 127) Then
                If Pointer_X < maxX Then
                    Ba_CLA(Pointer_X, Pointer_Y).Ansi = ByteArray(i)
                    Ba_CLA(Pointer_X, Pointer_Y).Color = FC_Byte
                    Ba_CLA(Pointer_X, Pointer_Y).BColor = BC_Byte
                    If cutX < Pointer_X Then cutX = Pointer_X   '取最大的x
                    Pointer_X = Pointer_X + 1
                End If
            

            ElseIf ByteArray(i) >= 128 Then '雙位元字
                If Pointer_X < maxX Then
                    tmpbiByte(1) = ByteArray(i)
                    preFC_Byte = FC_Byte
                    preBC_Byte = BC_Byte
                    biByteFlag = 1
                End If

            End If
            
        
        ElseIf CCMode = 1 Then
            '擷取彩色碼
            If ByteArray(i) >= 48 And ByteArray(i) <= 57 Then   '0~9
                tmpStr = tmpStr & Chr(ByteArray(i))
            ElseIf ByteArray(i) = 59 Then   ';
                tmpint2 = Val(tmpStr)
                tmpStr = ""
                If tmpint2 = 0 Then
                    colorResetFlag = 1
                ElseIf tmpint2 = 1 Then
                    rawFClight_Byte = 1
                ElseIf tmpint2 >= 40 And tmpint2 <= 48 Then
                    rawBC_Byte = tmpint2
                ElseIf tmpint2 >= 30 And tmpint2 <= 38 Then
                    rawFC_Byte = tmpint2
                End If
            ElseIf ByteArray(i) = 77 Or ByteArray(i) = 109 Then 'M or m
                tmpint2 = Val(tmpStr)
                tmpStr = ""
                If tmpint2 = 0 Then
                    colorResetFlag = 1
                ElseIf tmpint2 = 1 Then
                    rawFClight_Byte = 1
                ElseIf tmpint2 >= 40 And tmpint2 <= 48 Then
                    rawBC_Byte = tmpint2
                ElseIf tmpint2 >= 30 And tmpint2 <= 38 Then
                    rawFC_Byte = tmpint2
                End If
                If colorResetFlag = 1 Then  '初始化色彩
                    FC_Byte = 7
                    BC_Byte = 0
                    colorResetFlag = 0  '還原此標記
                End If
                If rawFC_Byte = 0 Then  '處理前景色彩
                    If rawFClight_Byte <> 0 Then
                        FC_Byte = (FC_Byte Mod 8) + 8
                        rawFClight_Byte = 0  '還原
                    End If
                Else
                    If rawFClight_Byte = 0 Then
                        If (FC_Byte \ 8) = 1 Then   '判斷有無繼承亮色
                            FC_Byte = GetFC(rawFC_Byte) + 8
                        Else
                            FC_Byte = GetFC(rawFC_Byte)
                        End If
                    Else
                        FC_Byte = GetFC(rawFC_Byte) + 8
                        rawFClight_Byte = 0 '還原
                    End If
                    rawFC_Byte = 0  '還原
                End If
                If rawBC_Byte <> 0 Then
                    BC_Byte = GetBC(rawBC_Byte)
                    rawBC_Byte = 0
                End If
                CCMode = 0
            End If
            
        'ElseIf CCMode = 2 Then
            
        End If

    Next i
    'ReDim Preserve Ba_CLA(0 To cutX, 0 To Pointer_Y) As ColorLayer
    If toClpB = 1 Then
        ReDim CPArr(cutX, Pointer_Y)
        For j = 0 To Pointer_Y
            For i = 0 To cutX
                CPArr(i, j) = Ba_CLA(i, j)
            Next i
        Next j
    Else
        Call ObjLibPo_Area(Ba_CLA, Arrf, X, Y, cutX, Pointer_Y)
    End If
    ReDim Ba_CLA(0)
End Function
