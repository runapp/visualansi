Attribute VB_Name = "Control_Collection"
Public Type cc_type
    ctl As Object
    type As CC_Type_Priority
    Value As Variant
End Type

Public Const C_Fore = 0
Public Const C_BG = 1
Public Const C_OBJ_deBG = 2
Public Const O_Pen_Text = 3

Public Enum CC_Type_Priority
    Chk = 0 'check box
    Opt = 1  'option
    
End Enum
Public Const CC_Check = 0
Public CC(10) As cc_type

Public Sub CC_Init()
    'Set CC(C_Fore).ctl = Form1.Check1
    'CC(C_Fore).type = CC_Check
    'CC_Update C_Fore
    
    Reg_C C_Fore, Form1.Check1, Chk   '前景
    Reg_C C_BG, Form1.Check2, Chk   '背景
    
    Reg_C C_OBJ_deBG, Form1.Check4, Chk   '物件 去背
    
    Reg_C O_Pen_Text, Form1.Option2, Opt   '背景
    
    
End Sub

Public Sub Reg_C(ByVal Index As Integer, ByRef ctl As Object, ByVal cc_type As CC_Type_Priority)
    Set CC(Index).ctl = ctl
    CC(Index).type = cc_type
    CC_Update Index
End Sub


Public Sub CC_Update(ByVal Index As Integer)
    
    
    Select Case CC(Index).type
        Case CC_Type_Priority.Chk, CC_Type_Priority.Opt
            CC(Index).Value = CC(Index).ctl.Value
            
        Case Else
            Debug.Print "cc_else"
    End Select
    
End Sub
