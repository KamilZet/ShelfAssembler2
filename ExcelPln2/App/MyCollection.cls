VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MyCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private someKeys As Dictionary
Private someCols As Collection
Public Function Add(o As Object, Key As String) As Object
    someKeys.Add Key, o
    someCols.Add o, Key
End Function
Public Property Get Count() As Long
    Count = someCols.Count
End Property
Public Property Get Item(vKey As Variant) As Object
    Set Item = someCols.Item(vKey)
End Property
Public Sub Remove(vKey As Variant)
    someKeys.Remove vKey
    someCols.Remove vKey
End Sub
Public Property Get NewEnum() As IUnknown
   Attribute NewEnum.VB_UserMemId = -4
   Set NewEnum = someCols.[_NewEnum]
End Property
Public Property Get Keys() As Variant
    Keys = someKeys.Keys
End Property
Private Sub Class_Initialize()
    Set someKeys = New Dictionary
    Set someCols = New Collection
End Sub

