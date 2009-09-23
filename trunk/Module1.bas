Attribute VB_Name = "Module1"
'定義圖畫基本元素
Type ColorLayer
    BColor As Byte
    Color As Byte
    Ansi As Integer
End Type

'定義目前檔案屬性變數模組
Type OpenedFilePropety
    FilePath As String
    Closed As Boolean
    filetype As Integer
    CurrentPage As Integer
    IsChanged As Boolean
End Type

Type BiAnsi
    First As Byte
    Second As Byte
End Type
Type IntS
    v As String * 1
End Type
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)

Public Type POINTAPI
        x As Long
        y As Long
End Type
'定義顏色交換功能的變數模組
Type ExChColorType
    CurrentSel As Byte '顏色交換功能_目前選取要選色的顏色方塊
    Color(0 To 3) As Byte

End Type

'定義前景來源為文章的變數模組
Type ForStringType
    str As String   '儲存目前的字串陣列
    StrLen(0 To 1) As Byte '儲存字串的長與寬
End Type

'定義程式中公用變數模組化
Public Type SysEnv
        PPA As Single  'pix per ansi
        Frontsize As Byte
        CheckSave As Byte
        HideSelect As Byte
        SSAnsi As Integer
        LastAnsi As Integer
        EdMode As Byte      '所選擇繪圖工具的模式
        ForeSource As Byte
        ForColor As Byte    '目前所選取的前景色
        BacColor As Byte    '目前所選取的背景色
        ExchC_Current As Byte
        ExChColor As ExChColorType  '顏色交換功能的變數模組
        ToolPBoxDown As Byte '將工具箱置底
        
        cDrawPos_X As Integer   '目前繪圖結束時的游標位置
        cDrawPos_Y As Integer
        InputFlag As Byte
        CopyingFlag As Byte '表示正在複製
        
        '拖曳移動
        Move_Draging As Boolean
        MD_X As Integer '拖曳位置基準
        MD_Y As Integer
        
End Type
'視窗位置API
Public Const HWND_TOPMOST = -1
Public Const HWND_TOP = 0
Public Const HWND_NOTOPMOST = -2
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_SHOWWINDOW = &H40
Public Declare Sub SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)

Public Declare Function Polygon Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long

Public Arrf() As ColorLayer '定義主要資料記憶體
Public timeLine() As Single '定義時間軸
Public CPArr() As ColorLayer   '複製貼上所用之記憶體

Public OFP As OpenedFilePropety '建立目前檔案屬性變數模組
Public SysInfo As SysEnv '建立程式公用變數模組
Public FString As ForStringType '建立前景為文章的變數(因長度不定會影響到設定檔 故獨立於公用環境變數)
Public Sub Main()
'Form9.Show
Form8.Show
Dim PauseTime, Start, Finish, TotalTime
    PauseTime = 1 / 86400   ' 設定暫停時間
    Start = Now    ' 設定開始暫停的時刻。
    Do While Now < Start + PauseTime
        DoEvents    ' 將程式執行權讓給其它程式。
    Loop
    If Command <> "" Then
        If Left(Command, 1) = "*" Then  '接到啟動方塊圖編輯器的命令
            Load Form14
        Else
            'Form1.MDIChild = True
            Load Form1
            'Form1.Show
        End If
    Else
        Load Form1
    End If
'Form7.Show
End Sub

Public Function ReadConfic(ByVal FilePosition As String) As SysEnv
    Dim ConficDS As SysEnv
    Dim ConficFile As Integer
    ConficFile = 3
    Open FilePosition For Binary As #ConficFile
    
    Get #ConficFile, 1, ConficDS
    Close ConficFile
    ReadConfic = ConficDS
End Function
Public Sub WriteConfic(ByVal FilePosition As String, ByRef ConficDS As SysEnv)
    Dim ConficFile As Integer
    ConficFile = 3
    Open FilePosition For Binary As #ConficFile
    
    Put #ConficFile, 1, ConficDS
    Close ConficFile
End Sub
Public Function Tlen(ByVal s As String) As Byte
    Tlen = LenB(StrConv(s, vbFromUnicode))
End Function
Public Function QBCToAnsiC(ByVal QBC As Integer) As String
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
    
    If QBC \ 8 = 1 Then
        AnsiC = "1;" & AnsiC
    Else
        AnsiC = ";" & AnsiC
    End If
    'If QBC = 15 Then AnsiC = "1"
    QBCToAnsiC = AnsiC
End Function
Public Function QBCToAnsiBC(ByVal QBC As Integer) As String
    Dim AnsiC As String
    
    Select Case QBC Mod 8
        Case Is = 0
            AnsiC = "40"
        Case Is = 1
            AnsiC = "44"
        Case Is = 2
            AnsiC = "42"
        Case Is = 3
            AnsiC = "46"
        Case Is = 4
            AnsiC = "41"
        Case Is = 5
            AnsiC = "45"
        Case Is = 6
            AnsiC = "43"
        Case Is = 7
            AnsiC = "47"
    End Select
    
    QBCToAnsiBC = AnsiC
End Function

