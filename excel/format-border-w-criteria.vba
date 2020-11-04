''' Usage: This should format a non-table grouping of cells for a "weekday" calendar
''' Imagine timestamps down the left column and days of the week in the header row

''' Example worksheet: '''
           ________ _________ ____________
'           Monday | Tuesday | Wendnesday |
' _________________________________________
' 9:00am  |________|_________|____________|
' 9:30am  |________|_________|____________|
' 10:00am |________|_________|____________|
' 10:30am |________|_________|____________|


Public Function UpdateFridayBorder(Optional StartingCellAddress As String = "B3")
On Error GoTo 0

''' Looks for columns with a Friday in the header section.
''' Formats the right-side border to help divide the weekday calendar into weekly chunks.

Dim RowStart As Long
Dim RowCount As Long
Dim ColStart As Long
Dim ColCount As Long
Dim c As Long

''' If we have a starting cell value arg, continue the function '''
If Not IsNull(StartingCellAddress) Or (StartingCellAddress & "") <> "" Then
    
    ''' Set starting row and column numeric values '''
    With Range(StartingCellAddress)
        RowStart = .Row
        ColStart = .Column
    
        ''' Get row and column counts '''
        ''' Offset to row and column with full set of values to avert any misreadings '''
        RowCount = Range(StartingCellAddress, .Offset(0, -1).End(xlDown)).Rows.Count
        ColCount = Range(StartingCellAddress, .Offset(-1, 0).End(xlToRight)).Columns.Count
    End With
    
    
    ''' Run the gambit from the start to second-to-last column. '''
    ''' (The last column in your grid will probably be formatted already.) '''
    For c = ColStart To ColCount
    
        ''' If the cell value is Friday, reformat the right-side border
        ''' to continuous and thick.
        If Cells(RowStart, c).Offset(-1, 0).Value = "Friday" Then
        
            ''' Update border down entire column '''
            With Range(Cells(RowStart, c).Address, Cells(RowStart + RowCount - 1, c).Address)
                .Borders(xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            
                ''' Border weight reference:
                '''     https://docs.microsoft.com/en-us/office/vba/api/excel.xlborderweight
                .Borders(xlEdgeRight).Weight = xlThick
            End With
            
        ''' Otherwise, revert cell border to continuous and thin. Just the way we like it! '''
        Else
            With Range(Cells(RowStart, c).Address, Cells(RowStart + RowCount - 1, c).Address)
                .Borders(xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            
                ''' Border weight! '''
                .Borders(xlEdgeRight).Weight = xlThin
            End With
        End If
    Next c
End If

End Function
