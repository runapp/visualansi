Attribute VB_Name = "RedoModule"
Public RD_Data As New Collection
Public RD_MaxNum As Integer
Public RD_CIndex As Integer '目前的位置
Public RD_MaxIndex As Integer   '可取消還原的位置
Public RD_Count As Integer  '可還原數量

Sub RD_Init()
    
    RD_MaxNum = 30
    RD_CIndex = 0
End Sub

Sub RD_Add(ByRef item As Variant)
    
End Sub

