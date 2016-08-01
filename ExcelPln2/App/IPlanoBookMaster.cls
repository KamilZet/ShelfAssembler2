VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IPlanoBookMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Function Initalize _
( _
    ByRef inWkbPath As String, _
    Optional ByRef inWorkspace As PlnWorkspace _
) As Boolean

End Function
        

Public Function ApplyChangesFrom(ByRef IPlanoBookSupplier) As Boolean

End Function

Public Function LoadPlanograms() As Boolean

End Function

Public Function IsPlanoValid(ByRef inWks As Excel.Worksheet) As Boolean

End Function

Public Function GetPlanoByFormat(ByRef inPlanoFormat As String) As Iplano

End Function

Public Property Get MetaSetting(ByRef inSetName As String) As Variant

End Property
