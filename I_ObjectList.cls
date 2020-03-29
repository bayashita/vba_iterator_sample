VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "I_ObjectList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


' Iterator で使用する処理のみ当インタフェースで定義する。

Public Property Get lastIndex() As Long
End Property

Public Function GetElementAt(index As Long) As Object
End Function

Public Function GetIterator() As C_Iterator
End Function

