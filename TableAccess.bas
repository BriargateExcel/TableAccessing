Attribute VB_Name = "TableAccess"
Option Explicit

Type DataAccessedType
    RowCount As Long
    ColumnCount As Long
    Data As Variant
End Type

Type ColumnDesignatorType
    ColumnName As String
    ColumnNumber As Long
End Type

Type ColumnFilterType
    ColumnName As String
    ColumnNumber As Long
    Operator As String
    Operand As String
End Type

Private Sub test()
        GetData Worksheets("Sheet1").ListObjects("ControlAccountTable"), , , "Control Account >= 8J6GM15223-02A"
    '    GetData Worksheets("Sheet1").ListObjects("ControlAccountTable"), 3, , "Control Account=8J6GM15223-02A"
    '    GetData Worksheets("Sheet1").ListObjects("ControlAccountTable"), 3, , "Control Account =8J6GM15223-02A"
    '    GetData Worksheets("Sheet1").ListObjects("ControlAccountTable"), 3, , "Control Account= 8J6GM15223-02A"
    '
    '    GetData Worksheets("Sheet1").ListObjects("ControlAccountTable"), 3, , "Control Account < 8J6GM15223-02A"
    '    GetData Worksheets("Sheet1").ListObjects("ControlAccountTable"), 3, , "Control Account<8J6GM15223-02A"
    '    GetData Worksheets("Sheet1").ListObjects("ControlAccountTable"), 3, , "Control Account <= 8J6GM15223-02A"
    '    GetData Worksheets("Sheet1").ListObjects("ControlAccountTable"), 3, , "Control Account <> 8J6GM15223-02A"
    '    GetData Worksheets("Sheet1").ListObjects("ControlAccountTable"), 3, , "Control Account<=8J6GM15223-02A"
    '    GetData Worksheets("Sheet1").ListObjects("ControlAccountTable"), 3, , "Control Account<>8J6GM15223-02A"
    '    GetData Worksheets("Sheet1").ListObjects("ControlAccountTable"), 3, , "Control Account <8J6GM15223-02A"
    '    GetData Worksheets("Sheet1").ListObjects("ControlAccountTable"), 3, , "Control Account <=8J6GM15223-02A"
    '    GetData Worksheets("Sheet1").ListObjects("ControlAccountTable"), 3, , "Control Account <>8J6GM15223-02A"
    '    GetData Worksheets("Sheet1").ListObjects("ControlAccountTable"), 3, , "Control Account< 8J6GM15223-02A"
    '    GetData Worksheets("Sheet1").ListObjects("ControlAccountTable"), 3, , "Control Account<= 8J6GM15223-02A"
    '    GetData Worksheets("Sheet1").ListObjects("ControlAccountTable"), 3, , "Control Account<> 8J6GM15223-02A"
    '
    '    GetData Worksheets("Sheet1").ListObjects("ControlAccountTable"), 3, , "Control Account > 8J6GM15223-02A"
    '    GetData Worksheets("Sheet1").ListObjects("ControlAccountTable"), 3, , "Control Account >= 8J6GM15223-02A"
    '
    '    GetData Worksheets("Sheet1").ListObjects("ControlAccountTable"), 3, , "Control Account>8J6GM15223-02A"
    '    GetData Worksheets("Sheet1").ListObjects("ControlAccountTable"), 3, , "Control Account>=8J6GM15223-02A"
    '
    '    GetData Worksheets("Sheet1").ListObjects("ControlAccountTable"), 3, , "Control Account >8J6GM15223-02A"
    '    GetData Worksheets("Sheet1").ListObjects("ControlAccountTable"), 3, , "Control Account >=8J6GM15223-02A"
    '
    '    GetData Worksheets("Sheet1").ListObjects("ControlAccountTable"), 3, , "Control Account> 8J6GM15223-02A"
    '    GetData Worksheets("Sheet1").ListObjects("ControlAccountTable"), 3, , "Control Account>= 8J6GM15223-02A"

End Sub

Public Function GetData( _
       ByVal SearchTable As ListObject, _
       Optional ByVal RowDesignator As Variant = "Empty", _
       Optional ByVal ColumnDesignator As String = "Empty", _
       Optional ByVal ColumnFilter As String = "Empty" _
       ) As DataAccessedType
    
    ' Future:
    ' Need provisions for multiple filters; "And" only; no "Or"
    ' ColumnFilter becomes a parameter array
    ' Alternately, use this routine for the first filter then past
    ' this routine's output to another routine for the next filter
    
    ' If RowDesignator is a number, that's the row to return
    ' If RowDesignator is "Empty" return the rows specified by ColumnFilter
    
    ' If ColumnDesignator has a value, that's the column to return
    ' If ColumnDesignator = "Empty", select the entire row
    '
    ' If ColumnFilter contains "=", there's only one row to return
    ' If ColumnFilter contains "<>", "<', ">", "<=", or ">="
    '   there are (potentially) multiple rows to return
    '
    ' Symbology for the table:
    ' RowDesignator (RD) can be numeric (N) or "Empty" (E)
    ' ColumnDesignator (CD) can be specified (S) or "Empty" (E)
    ' ColumnFilter (CF) can result in
    '   0 hits (0)
    '   1 hit (1)
    '   Multiple hits (M)
    '   "Empty" (E)
    '
    ' Table below:
    ' RD CD CF Result
    ' NS0 No data
    ' NS1 Single value
    ' NSM Single value
    ' NSE Single value
    ' NE0 No data
    ' NE1 One row
    ' NEM Single value
    ' NEE One row
    ' ES0 No data
    ' ES1 Single value
    ' ESM One column
    ' ESE One column
    ' EE0 No data
    ' EE1 Single value
    ' EEM Multiple rows, All columns
    ' EEE All Rows, All columns - the entire table
    '
    ' The output can be
    '   No data
    '       ColumnFilter evaluates to 0 rows (regardless of RowDesignator and ColumDesignator value)
    '   A single value
    '       RowDesignator is numeric and ColumnDesignator is specified and ColumnFilter evaluates to one row
    '       RowDesignator is numeric and ColumnDesignator is specified and ColumnFilter evaluates to multiple rows
    '       RowDesignator is numeric and ColumnDesignator is specified and ColumnFilter is empty
    '       RowDesignator is numeric and ColumnDesignator is "Empty" and ColumnFilter evaluates to multiple rows
    '       RowDesignator is "Empty" and ColumnDesignator is specified and ColumnFilter evaluates to a single row
    '   A single row
    '       RowDesignator is numeric and ColumnDesignator is "Empty" and ColumnFilter evaluates to one row
    '       RowDesignator is numeric and ColumnDesignator is "Empty" and ColumnFilter is empty
    '       RowDesignator is "Empty" and ColumnDesignator is "Empty" and ColumnFilter evaluates to one row
    '   A single column
    '       RowDesignator is "Empty" and ColumnDesignator is specified and ColumnFilter is "Empty"
    '       RowDesignator is "Empty" and ColumnDesignator is specified and ColumnFilter evaluates to multiple rows
    '   An array of rows and columns
    '       RowDesignator is "Empty" and ColumnDesignator is "Empty" and ColumnFilter evaluates to multiple rows
    '   All rows and columns
    '       RowDesignator is "Empty" and ColumnDesignator is "Empty" and ColumnFilter is "Empty"
    '
    ' Error messages:
    '   "Error Table" if the SearchTable is invalid
    '   "Error Row Designator" if Rowdesignator is invalid
    '   "Error Column Designator" if ColumnDesignator is invalid
    '   "Error Filter" if ColumnFilter is invalid
    '   "Error No Data" if ColumnFilter eliminates all the rows
    '   Note that the calling routine need only check for "Error"
    '       to determine if there's an error and need only
    '       go deeper if necessary
    
    ' Start of code
    '
    ' Verify that Searchtable is valid
    If ValidSearchTable(SearchTable) Then
        ' Valid Table
    Else
        ' Invalid Table
        GetData.Data = "Error Search Table"
        Exit Function
    End If
    
    ' Verify that RowDesignator is a valid row
    Dim RowNumber As Long
    Dim TempRowDesignator As Variant
    TempRowDesignator = ValidRowDesignator(SearchTable, RowDesignator)
    Select Case TempRowDesignator
    Case "Error"
        ' Invalid RowDesignator
        GetData.Data = "Error Row Designator"
        Exit Function
    Case "Empty"
        ' Empty RowDesignator
        RowNumber = 0
    Case Else
        ' RowDesignator must be a number
        RowNumber = TempRowDesignator
        GetData.RowCount = 1
    End Select
    
    ' Verify that ColumnDesignator is a valid column
    Dim ColumnName As String
    Dim ColumnNumber As Long
    Dim TempColumnDesignator As ColumnDesignatorType
    TempColumnDesignator = ValidColumnDesignator(SearchTable, ColumnDesignator)
    If TempColumnDesignator.ColumnName = "Error" Then
        ColumnName = "Error"
        ColumnNumber = 0
        Exit Function
    Else
        ' Valid ColumnDesignator
        ColumnName = TempColumnDesignator.ColumnName
        ColumnNumber = TempColumnDesignator.ColumnNumber
        GetData.ColumnCount = 1
    End If
    
    ' Verify that ColumnFilter is valid and parse it
    Dim FilterColumnName As String
    Dim FilterColumnNumber As Long
    Dim FilterOperator As String
    Dim FilterOperand As String
    Dim TempFilterValues As ColumnFilterType
    TempFilterValues = ValidFilter(SearchTable, ColumnFilter)
    If TempFilterValues.ColumnName = "Error" Then
        FilterColumnName = "Error"
        FilterColumnNumber = 0
        FilterOperator = vbNullString
        FilterOperand = vbNullString
        Exit Function
    Else
        ' Valid ColumnFilter
        FilterColumnName = TempFilterValues.ColumnName
        FilterColumnNumber = TempFilterValues.ColumnNumber
        FilterOperator = TempFilterValues.Operator
        FilterOperand = TempFilterValues.Operand
    End If
    
    If FilterColumnName <> "Error" Then
        Dim FilterCriteria As String
        FilterCriteria = FilterOperator & FilterOperand
        
        SearchTable.Range.AutoFilter Field:=FilterColumnNumber, Criteria1:=FilterCriteria
        
        Dim RowCount As Long
        On Error Resume Next
            RowCount = SearchTable.DataBodyRange.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count
            If Err.Number <> 0 Then RowCount = 0
        On Error GoTo 0
    Else
        ' Invalid ColumnFilter
        GetData.Data = "Error Column Filter"
    End If
    
    If RowDesignator <> "Empty" Then
        ' Valid RowDesignator
        If ColumnDesignator <> "Empty" Then
            ' Valid ColumnDesignator
            If ColumnFilter <> "Empty" Then
                ' Valid ColumnFilter
                Select Case RowCount
                Case 0 ' 1 row, 1 column, filter=0 rows; no data
                    GetData.Data = "Empty"
                    GetData.RowCount = 0
                    GetData.ColumnCount = 0
                Case 1 ' 1 row, 1 column, filter=1 row; one cell
                    GetData.Data = SearchTable.DataBodyRange(RowNumber, ColumnNumber)
                Case Else ' 1 row, 1 column, filter=multiple rows; one cell
                    GetData.Data = SearchTable.DataBodyRange(RowNumber, ColumnNumber)
                End Select
            Else ' 1 row, 1 column, empty filter; one cell
                GetData.Data = SearchTable.DataBodyRange(RowNumber, ColumnNumber)
            End If ' ColumnFilter <> "Empty"
        Else
            ' Empty ColumnDesignator
            GetData.ColumnCount = SearchTable.DataBodyRange.Columns.Count
            If ColumnFilter <> "Empty" Then
                ' Valid ColumnFilter
                Select Case RowCount
                Case 0 ' 1 row, unspecified columns, filter=0 rows; no data
                    GetData.Data = "Empty"
                Case 1 ' 1 row, unspecified columns, filter=1 row; one row
                    GetData.Data = SearchTable.DataBodyRange.Rows(RowNumber)
                Case Else ' 1 row, unspecified columns, filter=multiple rows; one row
                    GetData.Data = SearchTable.DataBodyRange.Rows(RowNumber)
                End Select
            Else ' 1 row, unspecified column, empty filter; entire row
                GetData.Data = SearchTable.DataBodyRange.Rows(RowNumber)
            End If ' ColumnFilter <> "Empty"
        End If ' ColumnDesignator <> "Empty"
    Else
        ' Empty RowDesignator
        If ColumnDesignator <> "Empty" Then
            ' Valid ColumnDesignator
            If ColumnFilter <> "Empty" Then
                ' Valid ColumnFilter
                Select Case RowCount
                Case 0 ' unspecified row, 1 column, filter=0 rows; no data
                    GetData.Data = "Empty"
                    GetData.RowCount = 0
                    GetData.ColumnCount = 0
                Case 1 ' unspecified row, 1 column, filter=1 row, one cell
                    GetData.Data = SearchTable.DataBodyRange.SpecialCells(xlCellTypeVisible).Columns(ColumnNumber)
                Case Else ' unspecified row, 1 column, filter=multiple rows; one column
                    GetData.Data = SearchTable.DataBodyRange.SpecialCells(xlCellTypeVisible).Columns(ColumnNumber)
                    GetData.RowCount = RowCount
                End Select
            Else ' empty row, one column, empty filter; one entire column
                GetData.Data = SearchTable.DataBodyRange.Columns(ColumnNumber)
            End If ' ColumnFilter <> "Empty"
        Else
            ' Empty ColumnDesignator
            GetData.ColumnCount = SearchTable.DataBodyRange.Columns.Count
            If ColumnFilter <> "Empty" Then
                ' Valid ColumnFilter
                Select Case RowCount
                Case 0 ' unspecified row, unspecified column, filter=0 rows; no data
                    GetData.Data = "Empty"
                    GetData.RowCount = 0
                    GetData.ColumnCount = 0
                Case 1 ' unspecified row, unspecified column, filter=1 row; one row, all columns
                    GetData.Data = SearchTable.DataBodyRange.SpecialCells(xlCellTypeVisible)
                Case Else ' unspecified row, unspecified column, filter=multiple rows; multiple rows, all columns
                    Dim TempRange As Range
                    Set TempRange = SearchTable.Range.SpecialCells(xlCellTypeVisible)
                    GetData.Data = SearchTable.DataBodyRange.SpecialCells(xlCellTypeVisible).Rows '.Cells '.AutoFilter '.Rows '   .Cells
                    GetData.RowCount = RowCount
                End Select
            Else ' empty row, empty column, empty filter; entire table
                GetData.Data = SearchTable.DataBodyRange
            End If ' ColumnFilter <> "Empty"
        End If ' ColumnDesignator <> "Empty"
    End If ' RowDesignator <> "Empty"
    
End Function

Private Function ValidSearchTable(ByVal SearchTable As ListObject) As Boolean
' Assumes SearchTable is a valid ListObject

    Dim CheckForValidSearchTable As String
    ValidSearchTable = True
    On Error GoTo 0
    CheckForValidSearchTable = SearchTable.Name
    If Err.Number <> 0 Then
        ValidSearchTable = False
        Exit Function
    End If
    On Error GoTo 0
End Function

Private Function ValidRowDesignator( _
        ByVal SearchTable As ListObject, _
        ByVal RowDesignator As Variant _
        ) As Variant

' Assumes SearchTable is a valid ListObject
    
    If RowDesignator = "Empty" Then
        ' Empty is a valid entry; means return an entire column
        ValidRowDesignator = 0
    Else
        Select Case VarType(RowDesignator)
        Case vbInteger, vbLong
            ' Verify that RowDesignator is in the range of the table's rows
            If RowDesignator >= 1 And RowDesignator <= SearchTable.Range.Rows.Count Then
                ValidRowDesignator = RowDesignator
            Else
                ' RowDesignator is out of range
                ValidRowDesignator = "Error"
            End If
        
        Case Else
            ' Erroneous RowDesignator data type
            ValidRowDesignator = "Error"
        End Select
    End If
    
End Function

Private Function ValidColumnDesignator( _
        ByVal SearchTable As ListObject, _
        ByVal ColumnDesignator As Variant _
        ) As ColumnDesignatorType
    
' Assumes SearchTable is a valid ListObject

    If ColumnDesignator = "Empty" Then
        ' "Empty" is a valid entry; means return an entire row
        ' Set ColumnNumber to 0 as a flag
        ValidColumnDesignator.ColumnName = "Empty"
        ValidColumnDesignator.ColumnNumber = 0
    Else
        Dim ColumnFound As Boolean
        ColumnFound = False
        
        Dim ColumnNumber As Long
        For ColumnNumber = 1 To SearchTable.ListColumns.Count
            If SearchTable.HeaderRowRange(, ColumnNumber) = ColumnDesignator Then
                ColumnFound = True
                ValidColumnDesignator.ColumnName = ColumnDesignator
                ValidColumnDesignator.ColumnNumber = ColumnNumber
                Exit For
            End If
        Next ColumnNumber
        If Not ColumnFound Then
            ValidColumnDesignator.ColumnName = "Error"
            ValidColumnDesignator.ColumnNumber = 0
        End If
    End If
End Function

Private Function ValidFilter( _
        ByVal SearchTable As ListObject, _
        ByVal ColumnFilter As String _
        ) As ColumnFilterType
    
' Assumes SearchTable is a valid ListObject

    If ColumnFilter = "Empty" Then
        ' "Empty" is a valid ColumnFilter value
        ValidFilter.ColumnName = "Empty"
        ValidFilter.ColumnNumber = 0
        ValidFilter.Operand = vbNullString
        ValidFilter.Operator = vbNullString
    Else
        ' Parse the ColumnFilter
        Dim I As Long
        Dim EndOfColumnName As Long
        Dim StartOfOperand As Long
        Dim PrevChar As String
        Dim ThisChar As String
        Dim NextChar As String
        For I = 1 To Len(ColumnFilter) - 1
            PrevChar = ThisChar
            ThisChar = Mid$(ColumnFilter, I, 1)
            NextChar = Mid$(ColumnFilter, I + 1, 1)
            
            Select Case ThisChar
            Case "="
                ValidFilter.Operator = "="
                
                If PrevChar = " " Then
                    ValidFilter.ColumnName = Mid$(ColumnFilter, 1, I - 2)
                Else
                    ValidFilter.ColumnName = Mid$(ColumnFilter, 1, I - 1)
                End If
                
                Dim TempValidColumnDesignator As ColumnDesignatorType
                TempValidColumnDesignator = ValidColumnDesignator(SearchTable, ValidFilter.ColumnName)
                If TempValidColumnDesignator.ColumnName = "Error" Then
                    ValidFilter.ColumnName = "Error"
                    Exit Function
                End If
                ValidFilter.ColumnNumber = TempValidColumnDesignator.ColumnNumber
                
                If NextChar = " " Then
                    ValidFilter.Operand = Mid$(ColumnFilter, I + 2, Len(ColumnFilter) - I - 1)
                Else
                    ValidFilter.Operand = Mid$(ColumnFilter, I + 1, Len(ColumnFilter) - I)
                End If
                
            Case "<"
                EndOfColumnName = I - 1
                While Mid$(ColumnFilter, EndOfColumnName, 1) = " "
                    EndOfColumnName = EndOfColumnName - 1
                Wend
                
                Select Case NextChar
                Case "="
                    ValidFilter.Operator = "<="
                    
                    StartOfOperand = I + 2
                    While Mid$(ColumnFilter, StartOfOperand, 1) = " "
                        StartOfOperand = StartOfOperand + 1
                    Wend
                    
                    ValidFilter.ColumnName = Mid$(ColumnFilter, 1, EndOfColumnName)
                
                    TempValidColumnDesignator = ValidColumnDesignator(SearchTable, ValidFilter.ColumnName)
                    If TempValidColumnDesignator.ColumnName = "Error" Then
                        ValidFilter.ColumnName = "Error"
                        Exit Function
                    End If
                    ValidFilter.ColumnNumber = TempValidColumnDesignator.ColumnNumber
                
                    ValidFilter.Operand = Mid$(ColumnFilter, StartOfOperand, Len(ColumnFilter) - StartOfOperand + 1)
                    Exit For
                Case ">"
                    ValidFilter.Operator = "<>"
                    
                    StartOfOperand = I + 2
                    While Mid$(ColumnFilter, StartOfOperand, 1) = " "
                        StartOfOperand = StartOfOperand + 1
                    Wend
                    
                    ValidFilter.ColumnName = Mid$(ColumnFilter, 1, EndOfColumnName)
                
                    TempValidColumnDesignator = ValidColumnDesignator(SearchTable, ValidFilter.ColumnName)
                    If TempValidColumnDesignator.ColumnName = "Error" Then
                        ValidFilter.ColumnName = "Error"
                        Exit Function
                    End If
                    ValidFilter.ColumnNumber = TempValidColumnDesignator.ColumnNumber
                
                    ValidFilter.Operand = Mid$(ColumnFilter, StartOfOperand, Len(ColumnFilter) - StartOfOperand + 1)
                    Exit For
                Case Else
                    ValidFilter.Operator = "<"
                    
                    StartOfOperand = I + 1
                    While Mid$(ColumnFilter, StartOfOperand, 1) = " "
                        StartOfOperand = StartOfOperand + 1
                    Wend
                    
                    ValidFilter.ColumnName = Mid$(ColumnFilter, 1, EndOfColumnName)
                
                    TempValidColumnDesignator = ValidColumnDesignator(SearchTable, ValidFilter.ColumnName)
                    If TempValidColumnDesignator.ColumnName = "Error" Then
                        ValidFilter.ColumnName = "Error"
                        Exit Function
                    End If
                    ValidFilter.ColumnNumber = TempValidColumnDesignator.ColumnNumber
                
                    ValidFilter.Operand = Mid$(ColumnFilter, StartOfOperand, Len(ColumnFilter) - StartOfOperand + 1)
                    Exit For
                End Select
            Case ">"
                EndOfColumnName = I - 1
                While Mid$(ColumnFilter, EndOfColumnName, 1) = " "
                    EndOfColumnName = EndOfColumnName - 1
                Wend
                
                If NextChar = "=" Then
                    ValidFilter.Operator = ">="
                    
                    StartOfOperand = I + 2
                    While Mid$(ColumnFilter, StartOfOperand, 1) = " "
                        StartOfOperand = StartOfOperand + 1
                    Wend
                    
                    ValidFilter.ColumnName = Mid$(ColumnFilter, 1, EndOfColumnName)
                
                    TempValidColumnDesignator = ValidColumnDesignator(SearchTable, ValidFilter.ColumnName)
                    If TempValidColumnDesignator.ColumnName = "Error" Then
                        ValidFilter.ColumnName = "Error"
                        Exit Function
                    End If
                    ValidFilter.ColumnNumber = TempValidColumnDesignator.ColumnNumber
                
                    ValidFilter.Operand = Mid$(ColumnFilter, StartOfOperand, Len(ColumnFilter) - StartOfOperand + 1)
                    Exit For
                Else
                    ValidFilter.Operator = ">"
                    
                    StartOfOperand = I + 1
                    While Mid$(ColumnFilter, StartOfOperand, 1) = " "
                        StartOfOperand = StartOfOperand + 1
                    Wend
                    
                    ValidFilter.ColumnName = Mid$(ColumnFilter, 1, EndOfColumnName)
                
                    TempValidColumnDesignator = ValidColumnDesignator(SearchTable, ValidFilter.ColumnName)
                    If TempValidColumnDesignator.ColumnName = "Error" Then
                        ValidFilter.ColumnName = "Error"
                        Exit Function
                    End If
                    ValidFilter.ColumnNumber = TempValidColumnDesignator.ColumnNumber
                
                    ValidFilter.Operand = Mid$(ColumnFilter, StartOfOperand, Len(ColumnFilter) - StartOfOperand + 1)
                    Exit For
                End If
            End Select
        Next I
    End If
    
End Function


