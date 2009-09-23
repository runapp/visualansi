VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDialogClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type
Public Enum FileFlags
    OFN_ALLOWMULTISELECT = &H200
    OFN_CREATEPROMPT = &H2000
    OFN_ENABLEHOOK = &H20
    OFN_ENABLETEMPLATE = &H40
    OFN_ENABLETEMPLATEHANDLE = &H80
    OFN_EXPLORER = &H80000 ' new look commdlg
    OFN_EXTENSIONDIFFERENT = &H400
    OFN_FILEMUSTEXIST = &H1000
    OFN_HIDEREADONLY = &H4
    OFN_LONGNAMES = &H200000 'force long names for 3.x modules
    OFN_NOCHANGEDIR = &H8
    OFN_NODEREFERENCELINKS = &H100000
    OFN_NOLONGNAMES = &H40000 ' force no long names for 4.x modules
    OFN_NONETWORKBUTTON = &H20000
    OFN_NOREADONLYRETURN = &H8000
    OFN_NOTESTFILECREATE = &H10000
    OFN_NOVALIDATE = &H100
    OFN_OVERWRITEPROMPT = &H2
    OFN_PATHMUSTEXIST = &H800
    OFN_READONLY = &H1
    OFN_SHAREAWARE = &H4000
    OFN_SHAREFALLTHROUGH = 2
    OFN_SHARENOWARN = 1
    OFN_SHAREWARN = 0
    OFN_SHOWHELP = &H10
    
    PD_ALLPAGES = &H0
    PD_COLLATE = &H10
    PD_DISABLEPRINTTOFILE = &H80000
    PD_ENABLEPRINTHOOK = &H1000
    PD_ENABLEPRINTTEMPLATE = &H4000
    PD_ENABLEPRINTTEMPLATEHANDLE = &H10000
    PD_ENABLESETUPHOOK = &H2000
    PD_ENABLESETUPTEMPLATE = &H8000
    PD_ENABLESETUPTEMPLATEHANDLE = &H20000
    PD_HIDEPRINTTOFILE = &H100000
    PD_NONETWORKBUTTON = &H200000
    PD_NOPAGENUMS = &H8
    PD_NOSELECTION = &H4
    PD_NOWARNING = &H80
    PD_PAGENUMS = &H2
    PD_PRINTSETUP = &H40
    PD_PRINTTOFILE = &H20
    PD_RETURNDC = &H100
    PD_RETURNDEFAULT = &H400
    PD_RETURNIC = &H200
    PD_SELECTION = &H1
    PD_SHOWHELP = &H800
    PD_USEDEVMODECOPIES = &H40000
    PD_USEDEVMODECOPIESANDCOLLATE = &H40000
End Enum
Private OFName As OPENFILENAME


Public Function ShowOpen() As Long
    OFName.lpstrFile = Space$(254)
    OFName.lpstrTitle = "開啟檔案"
    ShowOpen = GetOpenFileName(OFName)
End Function
Public Function ShowSave() As Long
    'OFName.lpstrFile = Space$(254)
    OFName.lpstrTitle = "儲存檔案"
    ShowSave = GetSaveFileName(OFName)
End Function

Private Sub Class_Initialize()

    
    OFName.lStructSize = Len(OFName)

    OFName.hInstance = App.hInstance

    OFName.lpstrFilter = "所有檔案 (*.*)" + Chr$(0) + "*.*" + Chr$(0)
    'create a buffer for the file
    OFName.lpstrFile = Space$(254)
    'set the maximum length of a returned file
    OFName.nMaxFile = 255

    OFName.lpstrFileTitle = Space$(254)

    OFName.nMaxFileTitle = 255

    OFName.lpstrInitialDir = App.Path

    OFName.lpstrTitle = "開啟檔案"
    'OFName.lpstrFile = "ddd.txt"
    OFName.flags = 0
End Sub

Public Sub Initialize(ByVal hwnd As Long)
    OFName.hwndOwner = hwnd
End Sub

Public Property Get DialogTitle() As String
    Title = OFName.lpstrTitle
End Property

Public Property Let DialogTitle(ByVal vNewValue As String)
    OFName.lpstrTitle = vNewValue
End Property
Public Property Get Filter() As String
    Title = OFName.lpstrFilter
End Property

Public Property Let Filter(ByVal vNewValue As String)
    OFName.lpstrFilter = Replace(vNewValue, "|", Chr(0)) & Chr(0)
End Property

Public Property Get FilePath() As String
    FilePath = Replace(Trim(OFName.lpstrFile), Chr(0), "")
End Property

Public Property Let FilePath(ByVal vNewValue As String)
    OFName.lpstrFile = vNewValue & Space$(254 - Tlen(vNewValue))
End Property

Public Property Get FileTitle() As Variant
    FileTitle = Replace(Trim(OFName.lpstrFileTitle), Chr(0), "")
End Property
Public Property Get FileName() As Variant
    FileName = Replace(Trim(OFName.lpstrFile), Chr(0), "")
End Property


