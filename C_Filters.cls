VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "C_Filters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Implements I_ObjectList

Private list() As C_Filter

' ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼　IObjectList共通実装　▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
Private privateIndex As Long

Private Sub Class_Initialize()
    privateIndex = -1
End Sub

Public Property Get I_ObjectList_lastIndex() As Long
    I_ObjectList_lastIndex = privateIndex
End Property

Public Function I_ObjectList_GetElementAt(index As Long) As Object
    Set I_ObjectList_GetElementAt = list(index)
End Function

Public Function I_ObjectList_GetIterator() As C_Iterator
    Dim iterator As C_Iterator
    Set iterator = New C_Iterator
    
    Call iterator.SetObjectList(Me)
    Set I_ObjectList_GetIterator = iterator
End Function

' ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲　IObjectList共通実装　▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲

Public Sub Add(object As C_Filter)
    privateIndex = privateIndex + 1
    ReDim Preserve list(privateIndex)
    Set list(privateIndex) = object
End Sub


