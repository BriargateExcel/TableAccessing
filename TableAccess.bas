Attribute VB_Name = "TableAccess"
Option Explicit

Private Const Module_Name As String = "TableAccess."

Private Const TempWorkSheetName As String = "TempWorkSheet"

Public Type TableType
    Headers As Variant ' The table's HeaderRowRange
    Body As Variant ' The table's DataBodyRage
    Valid As String ' Error message if TableType invalid
End Type

Private Type ColumnDesignatorType
    ColumnName As String
    ColumnNumber As Long
End Type
Public Function GetData( _
       SearchTable As TableType, _
       Optional ByVal RowDesignator As Variant = "Empty", _
       Optional ByVal ColumnDesignator As String = "Empty", _
       Optional ByVal ColumnFilter As String = "Empty" _
       ) As TableType
    
    ' ToDo:
    ' Need provisions for multiple filters; "And" only; no "Or"
    ' ColumnFilter becomes a parameter array
    ' Alternately, use this routine for the first filter then pass
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
    Const Routine_Name As String = Module_Name & "GetData"
    On Error GoTo ErrorHandler
    
    ' Verify that RowDesignator is valid
    Dim RowNumber As Long
    Dim TempRowDesignator As Variant
    TempRowDesignator = ValidRowDesignator(SearchTable, RowDesignator)
    Select Case Left(TempRowDesignator, 5)
    Case "Error"                                 ' Invalid RowDesignator
        GetData.Valid = "Error Row Designator"
        Exit Function
    Case "Empty"                                 ' Empty RowDesignator
        RowNumber = 0
    Case Else                                    ' Conclude that RowDesignator must be a number
        RowNumber = TempRowDesignator
    End Select
    
    ' Verify that ColumnDesignator is valid
    Dim ColumnName As String
    Dim ColumnNumber As Long
    Dim ColumnCount As Long
    Dim TempColumnDesignator As ColumnDesignatorType
    TempColumnDesignator = ValidColumnDesignator(SearchTable, ColumnDesignator)
    
    ColumnCount = 0
    Select Case Left(TempColumnDesignator.ColumnName, 5)
    Case "Error"                                 ' Invalid ColumnDesignator
        GetData.Valid = TempColumnDesignator.ColumnName
        Exit Function
    Case "Empty"                                 ' Empty ColumnDesignator
        ColumnName = "Empty"
        ColumnNumber = 0
        ColumnCount = UBound(SearchTable.Headers, 2)
    Case Else                                    ' Conclude that ColumnDesignator must be a valid column name
        ColumnName = ColumnDesignator
        ColumnNumber = TempColumnDesignator.ColumnNumber
        ColumnCount = UBound(SearchTable.Headers, 2)
    End Select
    
    ' Verify that ColumnFilter is valid and set up the SearchTable
    Dim ThisSearchTable As TableType
    ThisSearchTable = ValidFilter(SearchTable, ColumnFilter)
    
    Select Case Left(ThisSearchTable.Valid, 5)
    Case "Error"
        GetData.Valid = ThisSearchTable.Valid
        Exit Function
    Case "Empty"
        ThisSearchTable = SearchTable
    Case Else
        ' ThisSearchTable is already set up in the ValidFilter call
    End Select
    
    Dim RowCount As Long
    RowCount = UBound(ThisSearchTable.Body, 1)
    
    ' SearchTable, RowDesignator, ColumnDesignator, and ColumnFilter are all valid
    
    If RowDesignator <> "Empty" Then             ' Valid RowDesignator
        If ColumnDesignator <> "Empty" Then      ' Valid ColumnDesignator
            If ColumnFilter <> "Empty" Then      ' Valid ColumnFilter
                Select Case RowCount
                Case 0                           ' 1 row, 1 column, filter=0 rows; no data
                    GetData.Valid = "Empty"
                Case 1                           ' 1 row, 1 column, filter=1 row; one cell
                    GetData.Body = ThisSearchTable.Body(RowNumber, ColumnNumber)
                    GetData.Valid = "Valid"
                Case Else                        ' 1 row, 1 column, filter=multiple rows; one cell
                    GetData.Body = ThisSearchTable.Body(RowNumber, ColumnNumber)
                    GetData.Valid = "Valid"
                End Select
            Else                                 ' 1 row, 1 column, empty filter; one cell
                GetData.Body = ThisSearchTable.Body(RowNumber, ColumnNumber)
                GetData.Valid = "Valid"
            End If                               ' ColumnFilter <> "Empty"
        Else                                     ' Empty ColumnDesignator
            If ColumnFilter <> "Empty" Then
                ' Valid ColumnFilter
                Select Case RowCount
                Case 0                           ' 1 row, unspecified columns, filter=0 rows; no data
                    GetData.Valid = "Empty"
                Case 1                           ' 1 row, unspecified columns, filter=1 row; one row
                    GetData.Body = GetRow(ThisSearchTable, RowNumber)
                    GetData.Valid = "Valid"
                Case Else                        ' 1 row, unspecified columns, filter=multiple rows; one row
                    GetData.Body = GetRow(ThisSearchTable, RowNumber)
                    GetData.Valid = "Valid"
                End Select
            Else                                 ' 1 row, unspecified column, empty filter; entire row
                GetData.Body = GetRow(ThisSearchTable, RowNumber)
                GetData.Valid = "Valid"
            End If                               ' ColumnFilter <> "Empty"
        End If                                   ' ColumnDesignator <> "Empty"
    Else                                         ' Empty RowDesignator
        If ColumnDesignator <> "Empty" Then      ' Valid ColumnDesignator
            If ColumnFilter <> "Empty" Then      ' Valid ColumnFilter
                Select Case RowCount
                Case 0                           ' unspecified row, 1 column, filter=0 rows; no data
                    GetData.Valid = "Empty"
                Case 1                           ' unspecified row, 1 column, filter=1 row, one cell
                    GetData.Body = GetColumn(ThisSearchTable, ColumnNumber)
                    GetData.Valid = "Valid"
                Case Else                        ' unspecified row, 1 column, filter=multiple rows; one column
                    GetData.Body = GetColumn(ThisSearchTable, ColumnNumber)
                    GetData.Valid = "Valid"
                End Select
            Else                                 ' empty row, one column, empty filter; one entire column
                GetData.Body = GetColumn(ThisSearchTable, ColumnNumber)
                GetData.Valid = "Valid"
            End If                               ' ColumnFilter <> "Empty"
        Else                                     ' Empty ColumnDesignator
            If ColumnFilter <> "Empty" Then      ' Valid ColumnFilter
                Select Case RowCount
                Case 0                           ' unspecified row, unspecified column, filter=0 rows; no data
                    GetData.Valid = "Empty"
                Case 1                           ' unspecified row, unspecified column, filter=1 row; one row, all columns
                    GetData = ThisSearchTable
                Case Else                        ' unspecified row, unspecified column, filter=multiple rows; multiple rows, all columns
                    GetData = ThisSearchTable
                End Select
            Else                                 ' empty row, empty column, empty filter; entire table
                GetData = ThisSearchTable
            End If                               ' ColumnFilter <> "Empty"
        End If                                   ' ColumnDesignator <> "Empty"
    End If                                       ' RowDesignator <> "Empty"
    
    ' Delete the temporary worksheet
    Dim TempDisplayAlerts As Boolean
    TempDisplayAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False
    On Error Resume Next
    ActiveWorkbook.Worksheets(TempWorkSheetName).Delete
    On Error GoTo ErrorHandler
    Application.DisplayAlerts = TempDisplayAlerts
    
    '@Ignore LineLabelNotUsed
Done:
    Exit Function
ErrorHandler:
    RaiseError Err.Number, Err.Source, Routine_Name, Err.Description

End Function

Private Function ValidRowDesignator( _
        SearchTable As TableType, _
        ByVal RowDesignator As Variant _
        ) As Variant

    ' Assumes SearchTable is a valid ListObject
    
    Const Routine_Name As String = Module_Name & "ValidRowDesignator"
    On Error GoTo ErrorHandler
    
    If RowDesignator = "Empty" Then
        ' Empty is a valid entry; means return an entire column
        ValidRowDesignator = "Empty"
    Else
        Select Case VarType(RowDesignator)
        Case vbInteger, vbLong
            ' Verify that RowDesignator is in the range of the table's rows
            If RowDesignator >= 1 And RowDesignator <= UBound(SearchTable.Body, 1) Then
                ValidRowDesignator = RowDesignator
            Else
                ' RowDesignator is out of range
                ValidRowDesignator = "Error RowDesignator Out of Range"
            End If
        Case Else
            ' Erroneous RowDesignator data type
            ValidRowDesignator = "Error Bad Row Designator"
        End Select
    End If
    
    '@Ignore LineLabelNotUsed
Done:
    Exit Function
ErrorHandler:
    RaiseError Err.Number, Err.Source, Routine_Name, Err.Description

End Function

Private Function ValidColumnDesignator( _
        SearchTable As TableType, _
        ByVal ColumnDesignator As Variant _
        ) As ColumnDesignatorType
    
    ' Assumes SearchTable is a valid ListObject
    
    Const Routine_Name As String = Module_Name & "ValidColumnDesignator"
    On Error GoTo ErrorHandler
    
    If ColumnDesignator = "Empty" Then
        ' "Empty" is a valid entry; means return an entire row
        ' Set ColumnNumber to 0 as a flag
        ValidColumnDesignator.ColumnName = "Empty"
        ValidColumnDesignator.ColumnNumber = 0
    Else
        Dim ColumnFound As Boolean
        ColumnFound = False
        
        Dim ColumnNumber As Long
        For ColumnNumber = 1 To UBound(SearchTable.Headers, 2)
            If SearchTable.Headers(1, ColumnNumber) = ColumnDesignator Then
                ColumnFound = True
                ValidColumnDesignator.ColumnName = ColumnDesignator
                ValidColumnDesignator.ColumnNumber = ColumnNumber
                Exit For
            End If
        Next ColumnNumber
        If Not ColumnFound Then
            ValidColumnDesignator.ColumnName = "Error Column Name Not Found"
            ValidColumnDesignator.ColumnNumber = 0
        End If
    End If
    
    '@Ignore LineLabelNotUsed
Done:
    Exit Function
ErrorHandler:
    RaiseError Err.Number, Err.Source, Routine_Name, Err.Description

End Function

Private Function ValidFilter( _
        SearchTable As TableType, _
        ByVal ColumnFilter As String _
        ) As TableType
    
    Const Routine_Name As String = Module_Name & "ValidFilter"
    On Error GoTo ErrorHandler
    
    If ColumnFilter = "Empty" Then
        ' "Empty" is a valid ColumnFilter value
        ValidFilter.Valid = "Empty"
    Else
        ' Parse the ColumnFilter
        Dim I As Long
        Dim EndOfColumnName As Long
        Dim StartOfOperand As Long
        Dim PrevChar As String
        Dim ThisChar As String
        Dim NextChar As String
        Dim Operator As String
        Dim Operand As String
        Dim ColumnName As String
        Dim ColumnNumber As Long
        For I = 1 To Len(ColumnFilter) - 1
            PrevChar = ThisChar
            ThisChar = Mid$(ColumnFilter, I, 1)
            NextChar = Mid$(ColumnFilter, I + 1, 1)
            
            Select Case ThisChar
            Case "="
                Operator = "="
                
                If PrevChar = " " Then
                    ColumnName = Mid$(ColumnFilter, 1, I - 2)
                Else
                    ColumnName = Mid$(ColumnFilter, 1, I - 1)
                End If
                
                Dim TempValidColumnDesignator As ColumnDesignatorType
                TempValidColumnDesignator = ValidColumnDesignator(SearchTable, ColumnName)
                If TempValidColumnDesignator.ColumnName = "Error" Then
                    ValidFilter.Valid = "Error Bad Column Filter"
                    Exit Function
                End If
                ColumnNumber = TempValidColumnDesignator.ColumnNumber
                
                If NextChar = " " Then
                    Operand = Mid$(ColumnFilter, I + 2, Len(ColumnFilter) - I - 1)
                Else
                    Operand = Mid$(ColumnFilter, I + 1, Len(ColumnFilter) - I)
                End If
                
            Case "<"
                EndOfColumnName = I - 1
                While Mid$(ColumnFilter, EndOfColumnName, 1) = " "
                    EndOfColumnName = EndOfColumnName - 1
                Wend
                
                Select Case NextChar
                Case "="
                    Operator = "<="
                    
                    StartOfOperand = I + 2
                    While Mid$(ColumnFilter, StartOfOperand, 1) = " "
                        StartOfOperand = StartOfOperand + 1
                    Wend
                    
                    ColumnName = Mid$(ColumnFilter, 1, EndOfColumnName)
                
                    TempValidColumnDesignator = ValidColumnDesignator(SearchTable, ColumnName)
                    If TempValidColumnDesignator.ColumnName = "Error" Then
                        ValidFilter.Valid = "Error Bad Column Filter"
                        Exit Function
                    End If
                    ColumnNumber = TempValidColumnDesignator.ColumnNumber
                
                    Operand = Mid$(ColumnFilter, StartOfOperand, Len(ColumnFilter) - StartOfOperand + 1)
                    Exit For
                Case ">"
                    Operator = "<>"
                    
                    StartOfOperand = I + 2
                    While Mid$(ColumnFilter, StartOfOperand, 1) = " "
                        StartOfOperand = StartOfOperand + 1
                    Wend
                    
                    ColumnName = Mid$(ColumnFilter, 1, EndOfColumnName)
                
                    TempValidColumnDesignator = ValidColumnDesignator(SearchTable, ColumnName)
                    If TempValidColumnDesignator.ColumnName = "Error" Then
                        ValidFilter.Valid = "Error Bad Column Filter"
                        Exit Function
                    End If
                    ColumnNumber = TempValidColumnDesignator.ColumnNumber
                
                    Operand = Mid$(ColumnFilter, StartOfOperand, Len(ColumnFilter) - StartOfOperand + 1)
                    Exit For
                Case Else
                    Operator = "<"
                    
                    StartOfOperand = I + 1
                    While Mid$(ColumnFilter, StartOfOperand, 1) = " "
                        StartOfOperand = StartOfOperand + 1
                    Wend
                    
                    ColumnName = Mid$(ColumnFilter, 1, EndOfColumnName)
                
                    TempValidColumnDesignator = ValidColumnDesignator(SearchTable, ColumnName)
                    If TempValidColumnDesignator.ColumnName = "Error" Then
                        ColumnName = "Error"
                        Exit Function
                    End If
                    ColumnNumber = TempValidColumnDesignator.ColumnNumber
                
                    Operand = Mid$(ColumnFilter, StartOfOperand, Len(ColumnFilter) - StartOfOperand + 1)
                    Exit For
                End Select
            Case ">"
                EndOfColumnName = I - 1
                While Mid$(ColumnFilter, EndOfColumnName, 1) = " "
                    EndOfColumnName = EndOfColumnName - 1
                Wend
                
                If NextChar = "=" Then
                    Operator = ">="
                    
                    StartOfOperand = I + 2
                    While Mid$(ColumnFilter, StartOfOperand, 1) = " "
                        StartOfOperand = StartOfOperand + 1
                    Wend
                    
                    ColumnName = Mid$(ColumnFilter, 1, EndOfColumnName)
                
                    TempValidColumnDesignator = ValidColumnDesignator(SearchTable, ColumnName)
                    If TempValidColumnDesignator.ColumnName = "Error" Then
                        ValidFilter.Valid = "Error Bad Column Filter"
                        Exit Function
                    End If
                    ColumnNumber = TempValidColumnDesignator.ColumnNumber
                
                    Operand = Mid$(ColumnFilter, StartOfOperand, Len(ColumnFilter) - StartOfOperand + 1)
                    Exit For
                Else
                    Operator = ">"
                    
                    StartOfOperand = I + 1
                    While Mid$(ColumnFilter, StartOfOperand, 1) = " "
                        StartOfOperand = StartOfOperand + 1
                    Wend
                    
                    ColumnName = Mid$(ColumnFilter, 1, EndOfColumnName)
                
                    TempValidColumnDesignator = ValidColumnDesignator(SearchTable, ColumnName)
                    If TempValidColumnDesignator.ColumnName = "Error" Then
                        ValidFilter.Valid = "Error Bad Column Filter"
                        Exit Function
                    End If
                    ColumnNumber = TempValidColumnDesignator.ColumnNumber
                
                    Operand = Mid$(ColumnFilter, StartOfOperand, Len(ColumnFilter) - StartOfOperand + 1)
                    Exit For
                End If
            End Select
        Next I
        
        Dim FilterCriteria As String
        FilterCriteria = Operator & Operand
        ValidFilter = SetUpSearchTable(FilterCriteria, ColumnNumber, SearchTable)
    End If ' ColumnFilter = "Empty"
    
    '@Ignore LineLabelNotUsed
Done:
    Exit Function
ErrorHandler:
    RaiseError Err.Number, Err.Source, Routine_Name, Err.Description

End Function

Private Function SetUpSearchTable( _
        ByVal FilterCriteria As String, _
        ByVal FilterColumnNumber As Long, _
        SearchTable As TableType _
        ) As TableType
    
    Const Routine_Name As String = Module_Name & "GetFilteredData"
    On Error GoTo ErrorHandler
    
    
        
    '@Ignore LineLabelNotUsed
Done:
    Exit Function
ErrorHandler:
    RaiseError Err.Number, Err.Source, Routine_Name, Err.Description

End Function

Function GetRow( _
    SearchTable As TableType, _
    ByVal RowNum As Long _
    ) As Variant
    
    Const Routine_Name As String = Module_Name & "GetRow"
    On Error GoTo ErrorHandler
    
    ReDim Ary(UBound(SearchTable.Headers, 1), 1)
    
    Dim I As Long
    For I = 1 To UBound(SearchTable.Headers, 1)
        Ary(I, 1) = SearchTable.Body(RowNum, I)
    Next I
    
    GetRow = Ary
    
    '@Ignore LineLabelNotUsed
Done:
    Exit Function
ErrorHandler:
    RaiseError Err.Number, Err.Source, Routine_Name, Err.Description

End Function

Function GetColumn( _
    SearchTable As TableType, _
    ByVal ColumnNum As Long _
    ) As Variant
    
    Const Routine_Name As String = Module_Name & "GetColumn"
    On Error GoTo ErrorHandler
    
    ReDim Ary(UBound(SearchTable.Body, 1), 1)
    
    Dim I As Long
    For I = 1 To UBound(SearchTable.Body, 1)
        Ary(I, 1) = SearchTable.Body(I, ColumnNum)
    Next I
    
    GetColumn = Ary
    
    '@Ignore LineLabelNotUsed
Done:
    Exit Function
ErrorHandler:
    RaiseError Err.Number, Err.Source, Routine_Name, Err.Description

End Function

'Function GetAll( _
'    SearchTable As TableType _
'    ) As Variant
'
'    Const Routine_Name As String = Module_Name & "GetAll"
'    On Error GoTo ErrorHandler
'
'    Dim Ary(UBound(SearchTable.Body, 1), UBound(SearchTable.Body, 2))
'
'    Dim I As Long
'    Dim J As Long
'    For I = 1 To UBound(SearchTable.Body, 1)
'    For J = 1 To UBound(SearchTable.Headers, 1)
'            Ary(I, J) = SearchTable.Body(I, J)
'        Next J
'    Next I
'
'    GetRow = Ary
'
'    '@Ignore LineLabelNotUsed
'Done:
'    Exit Function
'ErrorHandler:
'    RaiseError Err.Number, Err.Source, Routine_Name, Err.Description
'
'End Function



