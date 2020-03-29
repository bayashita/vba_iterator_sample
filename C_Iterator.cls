VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "C_Iterator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


Private index As Long
Private list As I_ObjectList

Private Sub Class_Initialize()
    index = 0
End Sub


Public Function HasNext() As Boolean
    HasNext = (index <= list.lastIndex)
End Function


Public Function GetNext() As Object
    Set GetNext = list.GetElementAt(index)
    index = index + 1
End Function

Public Sub SetObjectList(objectList As I_ObjectList)
    Set list = objectList
End Sub
