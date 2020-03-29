VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "C_Datasource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


Private privateName As String

Public Property Get name() As String
    name = privateName
End Property

Public Property Let name(n As String)
    privateName = n
End Property

