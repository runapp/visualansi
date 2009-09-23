VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RegisterSystem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegQueryValue Lib "advapi32.dll" Alias "RegQueryValueA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal lpValue As String, lpcbValue As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, ByVal cbName As Long) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Private Declare Function RegSetValue Lib "advapi32.dll" Alias "RegSetValueA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
'最上層的hkey
Private Const HKEY_CLASSES_ROOT = &H80000000
Private Const HKEY_CURRENT_CONFIG = &H80000005
Private Const HKEY_CURRENT_USER = &H80000001
Private Const HKEY_DYN_DATA = &H80000006
Private Const HKEY_LOCAL_MACHINE = &H80000002
Private Const HKEY_PERFORMANCE_DATA = &H80000004
Private Const HKEY_USERS = &H80000003
'資料類型
Private Const REG_NONE = 0
Private Const REG_SZ = 1
Private Const REG_EXPAND_SZ = 2
Private Const REG_BINARY = 3
Private Const REG_DWORD = 4
Private Const REG_DWORD_BIG_ENDIAN = 5
Private Const REG_MULTI_SZ = 7


Public Function ReadReg(ByVal key As String, Optional ByVal KeyName As String)
'若無傳入KeyName則回傳subkey預設值
    Dim hKey As Long
    Dim ret As Long
    Dim lenData As Long
    Dim typeData As Long
    Dim hClassKey As Long
    Dim SubKey As String
    Dim KeyA() As String
    KeyA = Split(key, "\")
    
    Select Case KeyA(0)
        Case "HKEY_CLASSES_ROOT"
            hClassKey = HKEY_CLASSES_ROOT
        Case "HKEY_CURRENT_CONFIG"
            hClassKey = HKEY_CURRENT_CONFIG
        Case "HKEY_CURRENT_USER"
            hClassKey = HKEY_CURRENT_USER
        Case "HKEY_DYN_DATA"
            hClassKey = HKEY_DYN_DATA
        Case "HKEY_LOCAL_MACHINE"
            hClassKey = HKEY_LOCAL_MACHINE
        Case "HKEY_PERFORMANCE_DATA"
            hClassKey = HKEY_PERFORMANCE_DATA
        Case "HKEY_USERS"
            hClassKey = HKEY_USERS
        Case Else
            Debug.Print "error : keyclass not found"
            Exit Function
    End Select
    SubKey = KeyA(1)
    If UBound(KeyA) > 2 Then
        For i = 2 To UBound(KeyA)
            SubKey = SubKey & "\" & KeyA(i)
        Next i
    End If
    
    Debug.Print "keya(0)=" & KeyA(0)
    Debug.Print "subkey=" & SubKey
    Debug.Print "keyname=" & KeyName
    ret = RegOpenKey(hClassKey, SubKey, hKey)
    If ret <> 0 Then '表失敗
        
        Debug.Print "ret=" & ret & "error : OPENKEY error maybe not exit"
        Exit Function
    End If
    
    ret = RegQueryValueEx(hKey, KeyName, 0, typeData, ByVal 0, lenData)
    
    If ret <> 0 Then '表失敗
        Debug.Print "error : RegQueryValueEx error "
        Exit Function
    End If
    
    Select Case typeData
        Case REG_SZ, REG_EXPAND_SZ, REG_MULTI_SZ
            Dim s As String
            s = String(lenData, Chr(0))
            RegQueryValueEx hKey, KeyName, ByVal 0, typeData, ByVal s, lenData
            If typeData = REG_SZ Or typeData = REG_EXPAND_SZ Then
                s = Left(s, InStr(s, Chr(0)) - 1)
            End If
            ReadReg = s
        Case REG_DWORD, REG_DWORD_BIG_ENDIAN
            Dim L As Long
            RegQueryValueEx hKey, KeyName, ByVal 0, typeData, L, lenData
            ReadReg = L
        Case REG_BINARY
            Dim bArr() As Byte
            ReDim bArr(0 To lenData - 1) As Byte
            RegQueryValueEx hKey, KeyName, ByVal 0, typeData, bArr(0), lenData
            ReadReg = bArr
    End Select
    
End Function


Public Function WriteReg(ByVal KeyStr As String, ByVal KeyName As String, ByVal KeyValue, Optional KeyType As String)
'typename

    Dim hKey As Long
    Dim ret As Long
    Dim lenData As Long
    Dim typeData As Long
    Dim hClassKey As Long
    Dim SubKey As String
    Dim KeyA() As String
    KeyA = Split(KeyStr, "\")
    If KeyType = "" Then
        KeyType = TypeName(KeyValue)
    End If
    Select Case KeyA(0)
        Case "HKEY_CLASSES_ROOT"
            hClassKey = HKEY_CLASSES_ROOT
        Case "HKEY_CURRENT_CONFIG"
            hClassKey = HKEY_CURRENT_CONFIG
        Case "HKEY_CURRENT_USER"
            hClassKey = HKEY_CURRENT_USER
        Case "HKEY_DYN_DATA"
            hClassKey = HKEY_DYN_DATA
        Case "HKEY_LOCAL_MACHINE"
            hClassKey = HKEY_LOCAL_MACHINE
        Case "HKEY_PERFORMANCE_DATA"
            hClassKey = HKEY_PERFORMANCE_DATA
        Case "HKEY_USERS"
            hClassKey = HKEY_USERS
        Case Else
            Debug.Print "error : keyclass not found"
            Exit Function
    End Select
    SubKey = KeyA(1)
    If UBound(KeyA) > 2 Then
        For i = 2 To UBound(KeyA)
            SubKey = SubKey & "\" & KeyA(i)
        Next i
    End If
    
    Debug.Print "keya(0)=" & KeyA(0)
    Debug.Print "subkey=" & SubKey
    Debug.Print "keyname=" & KeyName
    ret = RegOpenKey(hClassKey, SubKey, hKey)
    If ret <> 0 Then '表失敗
        
        Debug.Print "ret=" & ret & "error : OPENKEY error maybe not exit"
        Exit Function
    End If
    Select Case KeyType
    
        Case "string", "str", "REG_SZ"
            lenData = Tlen(KeyValue)
            RegSetValueEx hKey, KeyName, 0, REG_SZ, ByVal KeyValue, lenData
            WriteReg = 1
        Case "REG_EXPAND_SZ"
            lenData = Tlen(KeyValue)
            RegSetValueEx hKey, KeyName, 0, REG_EXPAND_SZ, ByVal KeyValue, lenData
            WriteReg = 1
        Case "REG_MULTI_SZ"
            lenData = Tlen(KeyValue)
            RegSetValueEx hKey, KeyName, 0, REG_MULTI_SZ, ByVal KeyValue, lenData
            WriteReg = 1
        
        Case "long", "integer", "REG_DWORD"
            Dim tempL As Long
            tempL = KeyValue
            RegSetValueEx hKey, KeyName, 0, REG_DWORD, tempL, 4
            WriteReg = 1
        
        Case "REG_DWORD_BIG_ENDIAN"
            Dim tempL1 As Long
            tempL = KeyValue
            RegSetValueEx hKey, KeyName, 0, REG_DWORD_BIG_ENDIAN, tempL1, 4
            WriteReg = 1
        
        Case "byte()", REG_BINARY, "byte"
            lenData = UBound(KeyValue)
            RegSetValueEx hKey, KeyName, 0, REG_BINARY, KeyValue, lenData
            WriteReg = 1
        Case Else
            WriteReg = 0
    End Select
    Exit Function
out:
    Debug.Print "WriteReg Error Out "
    
End Function
    
Private Function Tlen(ByVal str As String) As Long
    
    Tlen = LenB(StrConv(str, vbFromUnicode))
    
End Function

Public Function CreatRegKey(ByVal KeyStr As String)
    
End Function

Public Sub DelReg(ByVal KeyStr As String)

End Sub
