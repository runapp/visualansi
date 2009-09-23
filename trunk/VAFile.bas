Attribute VB_Name = "VAFile"
Public Type VAFileInfo
    IDC As Integer
    'IDC 識別碼
    filetype As Integer
    '檔案類型
    '1 單頁畫
    '2 多頁畫
    '3 動畫
    '4 方塊畫
    '5 方塊動畫
    ArrXUbound As Integer
    ArrYUbound As Integer
    ArrZLenth As Integer
    '檔案資訊的結構化資料
End Type


Public Sub VA_SaveFile(ByVal filepath As String, ByRef filedata() As ColorLayer, ByRef Fp As OpenedFilePropety, ByRef intimeLine() As Single)
On Error GoTo out
    Dim VAFI As VAFileInfo
    Dim OpFile As Integer
    '寫入檔案資訊
    VAFI.IDC = 828
    VAFI.filetype = Fp.filetype
    VAFI.ArrXUbound = UBound(filedata, 1)
    VAFI.ArrYUbound = UBound(filedata, 2)
    If VAFI.filetype = 2 Or VAFI.filetype = 3 Then
        VAFI.ArrZLenth = UBound(filedata, 3)
    End If
    
    'Debug.Print "VAFI.ArrXUbound =" & UBound(filedata, 1)
    'Debug.Print "VAFI.ArrYUbound =" & UBound(filedata, 2)
    OpFile = 1
    
    If Form1.FileSys.FileExists(filepath) = True Then
        '以binary開啟檔案 要砍掉原來的 才不會留下之前的資訊
       
        Kill filepath
        
        '若檔案不存在 kill會發生錯誤
    End If
        
    Open filepath For Binary As #OpFile
    Put #OpFile, 1, VAFI
    Put #OpFile, 11, filedata
    Put #OpFile, , intimeLine
    
    Debug.Print "savefile  " & filepath
    Close #OpFile
    Exit Sub
out:

End Sub

Public Sub VA_ReadFile(ByVal filepath As String, ByRef filedata() As ColorLayer, ByRef intimeLine() As Single)
    
    Dim VAFI As VAFileInfo
    Dim OpFile As Integer
    Dim tmpLen As Long
    
    OpFile = 1
    
    Open filepath For Binary As #OpFile
    Get #OpFile, 1, VAFI
    'ReDim filedata(VAFI.ArrXUbound, VAFI.ArrYUbound) As ColorLayer
    Select Case VAFI.filetype
        Case Is = 1
            Call Form1.SetSize(VAFI.ArrXUbound + 1, VAFI.ArrYUbound + 1, 1, 1)
        Case Is = 2
            Call Form1.SetSize(VAFI.ArrXUbound + 1, VAFI.ArrYUbound + 1, VAFI.ArrZLenth, 2)
        Case Is = 3
            Call Form1.SetSize(VAFI.ArrXUbound + 1, VAFI.ArrYUbound + 1, VAFI.ArrZLenth, 3)
    End Select
    
    Get #OpFile, 11, filedata
    tmpLen = 4 * (UBound(filedata, 1) + 1) * (UBound(filedata, 2) + 1) * UBound(filedata, 3) + 11 '計算前面資料的長度
    If tmpLen < LOF(OpFile) Then
        Debug.Print "讀取時間軸" & tmpLen & "<" & LOF(OpFile)
        Get #OpFile, tmpLen, timeLine
    Else
        'Get #OpFile, tmpLen, timeLine
        Debug.Print "沒時間軸" & tmpLen & "<" & LOF(OpFile)
    End If
    
    
    Close #OpFile


End Sub

Public Sub OpenNewFile()
Form4.Show vbModal
End Sub

Public Sub CreatNewFile(ByRef filedata() As ColorLayer, VAFI As VAFileInfo)
Debug.Print "建立一個新檔案 檔案格式" & VAFI.filetype & " 寬: " & VAFI.ArrXUbound & " 長: " & VAFI.ArrYUbound
Select Case VAFI.filetype

    Case Is = 1
        Call Form1.SetSize(VAFI.ArrXUbound + 1, VAFI.ArrYUbound + 1, 1, 1)
    Case Is = 2
        Call Form1.SetSize(VAFI.ArrXUbound + 1, VAFI.ArrYUbound + 1, VAFI.ArrZLenth, 2)
    Case Is = 3
        Call Form1.SetSize(VAFI.ArrXUbound + 1, VAFI.ArrYUbound + 1, VAFI.ArrZLenth, 3)
End Select

End Sub


