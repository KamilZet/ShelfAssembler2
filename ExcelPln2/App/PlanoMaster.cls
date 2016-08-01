VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PlanoMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Implements Iplano


Private format_     As String
Private posRange_   As Range
'
Private sourceWks_      As Excel.Worksheet
Private positions_()    As Collection

Private startCellAddress_ As String
Private Const validPosRgb_ = 6750207 'yellow?
Private isInitialized_  As Boolean
Private rowsCount_      As Long
Private colsCount_      As Long
'

Private Property Get Iplano_Format() As String
    Iplano_Format = format_
End Property

Private Property Get Iplano_GetPositionByRowCol(inRow As Integer, inCol As Integer) As Variant

End Property

Private Function Iplano_IsPositionValid(inCell As Range) As Boolean
    With inCell
        Iplano_IsPositionValid = (.Value <> 1 And _
                                  .Interior.Color = validPosRgb_)
    End With
End Function

Public Function LoadPositions() As Boolean
    LoadPositions = Iplano_LoadPositions
End Function

Private Function Iplano_LoadPositions() As Boolean
    
On Error GoTo ErrorExit
    
    Dim plnRange As Range:      Set plnRange = Iplano_PositionsRange
    
    For i = 1 To (plnRange.Cells.Count)
        If (Iplano_IsPositionValid(plnRange.Cells(i))) Then
            'Set tmpPosition = LoadPosition(cell)
            ReDim Preserve positions_(0 To j)
            Set positions_(j) = Iplano_LoadPosition(plnRange.Cells(i)) 'tmpPosition
            ' if cell behind the position is merged, then skip number of columns merged into it!
            j = j + 1
        End If
        i = i + (plnRange.Cells(i).MergeArea.Columns.Count - 1)
    Next i

NormalExit:
Exit Function

ErrorExit:
    Iplano_LoadPositions = False
    
End Function

Private Function Iplano_Merge(ByRef inPln As Iplano) As Boolean
    If Not (isInitialized_) Then
        
End Function

Public Function Initialize(ByRef inWks As Excel.Worksheet) As Boolean
    IPlano_Initialize inWks
End Function

Private Function IPlano_Initialize(ByRef SourceWorksheet As Excel.Worksheet) As Boolean
        
On Error GoTo ErrorExit
        
        
    Set sourceWks_ = SourceWorksheet
    
    startCellAddress_ = findGridStartCell.Address
    
    'todo
    'planogram format value is stored in a merged cells, which makes look up procedure hard to implement w/o loops...
    'assumption is made that all planograms format values are stored in cell "C3"
    format_ = SourceWorksheet.Cells(3, 3)
'            findPlanogramFormat (sourceWorksheet) 'depreceated method, possibly could be re-implemented???
    
    #If (Development) Then
        Debug.Print "planobook " & SourceWorksheet.Cells(3, 3) & " position load started at " & Time
    #End If
    
    
    'Iplano_LoadPositions
    LoadPositionsDev
    
    #If (Development) Then
        Debug.Print "planobook " & SourceWorksheet.Cells(3, 3) & " positioni load finished at " & Time
    #End If
    
NormalExit:
    isInitialized_ = True
Exit Function


ErrorExit:

    
End Function

Public Property Get Format() As String
    
    Format = format_

End Property

Private Function Iplano_LoadPosition(ByRef inCell As Excel.Range) As position
        
    Dim tmpPosition As position
    Set tmpPosition = New position
    
    Dim tmpProdNames As String: tmpProdNames = Split(inCell.Value, Globals.ProdNameSep)
    Dim tmpImages()  As Excel.Shape

    Dim pic As Excel.Shape
    For i = 1 To sourceWks_.Shapes.Count
        If (inCell.Address = sourceWks_.Shapes(i).TopLeftCell.Address) Then
            tmpPosition.AddImage sourceWks_.Shapes(i)
            Exit For
        End If
    Next i
        
    With tmpPosition
        
        .Name = inCell.Value
        .StartRow = inCell.Row - Iplano_PositionsRangeStartCell.Row + 1
        .StartColumn = inCell.Column - Iplano_PositionsRangeStartCell.Column + 1
        If (inCell.MergeCells) Then
            .Rows = inCell.MergeArea.Rows.Count
            .Cols = inCell.MergeArea.Columns.Count
        End If
    End With
    
    Set Iplano_LoadPosition = tmpPosition

End Function

Private Property Get Iplano_GetPositions() As position()
    Iplano_GetPositions = positions_
End Property

Private Property Get Iplano_PositionsRange() As Range
    Set Iplano_PositionsRange = Iplano_PositionsRangeStartCell.Resize(Iplano_RowsCount, Iplano_ColsCount)
End Property


Private Property Get Iplano_PositionsRangeStartCell() As Range
    Set Iplano_PositionsRangeStartCell = sourceWks_.Range(startCellAddress_)
End Property

Public Function colsCount() As Integer
    colsCount = Iplano_PositionsRangeStartCell.Offset(1, -1).End(xlDown).Row - Iplano_PositionsRangeStartCell.Row
End Function

Private Property Get Iplano_ColsCount() As Integer
    Iplano_ColsCount = Iplano_PositionsRangeStartCell.Offset(-1, 0).End(xlToRight).Column - Iplano_PositionsRangeStartCell.Column + 1
End Property

Private Property Get Iplano_RowsCount() As Integer
    Iplano_RowsCount = Iplano_PositionsRangeStartCell.Offset(0, -1).End(xlDown).Row - Iplano_PositionsRangeStartCell.Row + 1
End Property

Public Property Get IsInitialized() As Boolean
    IsInitialized = (isInitialized_ And Helper.IsArrayAllocated(positions_))
End Property

Public Function LoadPositionsDev() As Boolean
    
On Error GoTo ErrorExit
    
    Dim c As New Collection
    Dim p As position
    Dim sh As Shape
    Dim newPosId As String 'key used to uniquely identify a position within a positions range
    
    For Each sh In sourceWks_.Shapes
        
        'generate mock position id
        newPosId = sh.TopLeftCell.Row & RowColSep & sh.TopLeftCell.Column
        
        If (Contains(c, newPosId)) Then
            Set p = c(newPosId)
            p.AddImage sh
        Else
            Set p = New position
            p.Initialize sh.TopLeftCell, Me
            p.AddImage sh
            c.Add key:=newPosId, Item:=p
        End If
        Set p = Nothing
    Next sh
    
'    Stop

'Each range not included

ErrorExit:
Exit Function


End Function

Private Property Get Iplano_SourceWorksheet() As Worksheet
    Set SourceWorksheet = sourceWks_
End Property

Private Property Get Iplano_VendorRange(inVendorName As String) As Range
    
    Set Iplano_VendorRange = CollectVendorByColor(inVendorName)

End Property

Private Function CollectVendorByColor(ByRef vendName As String) As Range
    
    Dim allCells As Range
    Set allCells = Iplano_PositionsRange
    
    Dim vendColor As Variant
    vendColor = SUPPLIERCONFIG.VendorColor(vendName)
    
    Dim cell As Range
    Dim Result As Range
    Set Result = allCells.Cells(1)
    
    For Each cell In allCells
        If (cell.Interior.Color = vendColor) Then
            Set Result = Union(Result, cell)
            Debug.Print cell.Address
        End If
    Next cell
    
    Set CollectVendorByColor = Result
    
End Function

Private Function findGridStartCell() As Range
    Set findGridStartCell = sourceWks_.Cells.Find( _
                                                    what:="1", _
                                                    lookat:=xlWhole, _
                                                    searchorder:=xlByColumns) _
                                                .Offset(-1, 1)
    Do Until findGridStartCell.Value = "1"
        Set findGridStartCell = ActiveSheet.Cells.FindNext
    Loop
    
    Set findGridStartCell = findGridStartCell.Offset(rowoffset:=1)
    
End Function

Private Function findPlanogramFormat(ByRef SourceWorksheet As Worksheet) As String

'On Error Resume Next

findPlanogramFormat = SourceWorksheet.Cells.Find( _
                                    what:="PLANOGRAM", _
                                    MatchCase:=True, _
                                    lookat:=xlWhole, _
                                    searchorder:=xlByColumns) _
                                .Offset(rowoffset:=1).Value

If (err) Then findPlanogramFormat = ERROR_NO_PLN_FORMAT_FOUND

End Function
