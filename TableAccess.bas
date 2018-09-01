Attribute VB_Name = "TableAccess"
Option Explicit

Private Const Module_Name As String = "TableAccess."

Public Type TableType
    ' Headers
    '   The table's HeaderRowRange
    '   2D array
    '   Only has row 1 (first parameter)
    '   The column names are in the column (second) parameter of the array
    '   If it exists, Row 0 is not used
    ' Body
    '   The table's DataBodyRage
    '   Each table row is designated by the first (row) parameter of the array
    '   Each column is designated by the second (column) parameter of the array
    '   If it exists, Row 0 is not used
    ' Valid
    '   "Valid" if TableType valid
    '   Error message if TableType invalid
    Headers As Variant
    Body As Variant
    Valid As String
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
       
    ' Purpose
    '   Return the subset of SearchTable as specified in the other parameters
    ' Assumptions
    
    ' Future:
    ' Add provisions for multiple filters; "And" only; no "Or"
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
    ' NS0 Makes no sense
    ' NS1 Makes no sense
    ' NSM Makes no sense
    ' NSE Single value
    ' NE0 Makes no sense
    ' NE1 Makes no sense
    ' NEM Makes no sense
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
    '       RowDesignator is numeric and ColumnDesignator is specified and ColumnFilter is empty
    '   A single row
    '       RowDesignator is numeric and ColumnDesignator is "Empty" and ColumnFilter is empty
    '       RowDesignator is "Empty" and ColumnDesignator is specified and ColumnFilter evaluates to a single row
    '       RowDesignator is "Empty" and ColumnDesignator is "Empty" and ColumnFilter evaluates to a single row
    '   A single column
    '       RowDesignator is "Empty" and ColumnDesignator is specified and ColumnFilter is "Empty"
    '       RowDesignator is "Empty" and ColumnDesignator is specified and ColumnFilter evaluates to multiple rows
    '   An array of rows and columns
    '       RowDesignator is "Empty" and ColumnDesignator is "Empty" and ColumnFilter evaluates to multiple rows
    '   All rows and columns
    '       RowDesignator is "Empty" and ColumnDesignator is "Empty" and ColumnFilter is "Empty"
    '   Makes no sense to specify a row and a filter
    '       RowDesignator is numeric and ColumnFilter <> "Empty"
    '
    ' Error messages:
    '   "Error Table" if the SearchTable is invalid
    '   "Error Row Designator" if Rowdesignator is invalid
    '   "Error RowDesignator Out of Range"
    '   "Error Column Designator" if ColumnDesignator is invalid
    '   "Error Column Name Not Found"
    '   "Error Filter" if ColumnFilter is invalid
    '   "Error No Data" if ColumnFilter eliminates all the rows
    '   "Error Can't have a specific row and a column filter"
    '   Note that the calling routine need only check for "Error"
    '       to determine if there's an error and need only
    '       go deeper if necessary
    
    ' Start of code
    '
    Const Routine_Name As String = Module_Name & "GetData"
    On Error GoTo ErrorHandler
    
    ' Verify that SearchTable is valid
    If Left$(SearchTable.Valid, 5) = "Error" Then
        GetData.Valid = "Error Table"
        Exit Function
    End If
    
    ' Verify that RowDesignator is valid
    Dim RowNumber As Long
    Dim TempRowDesignator As Variant
    TempRowDesignator = ValidRowDesignator(SearchTable, RowDesignator)
    Select Case Left$(TempRowDesignator, 5)
    Case "Error"                                 ' Invalid RowDesignator
        GetData.Valid = "Error Row Designator"
        Exit Function
    Case "Empty"                                 ' Empty RowDesignator
        RowNumber = 0
    Case Else                                    ' Conclude that RowDesignator must be a number
        If ColumnFilter <> "Empty" Then
            GetData.Valid = "Error Can't have a specific row and a column filter"
            Exit Function
        End If
        RowNumber = TempRowDesignator
    End Select
    
    ' Verify that ColumnDesignator is valid
    Dim ColumnName As String
    Dim ColumnNumber As Long
    Dim ColumnCount As Long
    Dim TempColumnDesignator As ColumnDesignatorType
    TempColumnDesignator = ValidColumnDesignator(SearchTable, ColumnDesignator)
    
    ColumnCount = 0
    Select Case Left$(TempColumnDesignator.ColumnName, 5)
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
    
    Select Case Left$(ThisSearchTable.Valid, 5)
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
                ' The case where RowDesignator <> "Empty" and ColumnFilter <> "Empty" cannot exist
                GetData.Valid = "Error Can't have a specific row and a column filter"
            Else                                 ' 1 row, 1 column, empty filter; one cell
                GetData.Body = ThisSearchTable.Body(RowNumber, ColumnNumber)
                GetData.Valid = "Valid"
            End If                               ' ColumnFilter <> "Empty"
        Else                                     ' Empty ColumnDesignator
            If ColumnFilter <> "Empty" Then
                ' The case where RowDesignator <> "Empty" and ColumnFilter <> "Empty" cannot exist
            Else                                 ' 1 row, unspecified column, empty filter; entire row
                GetData.Body = GetRow(ThisSearchTable, RowDesignator)
                GetData.Valid = "Valid"
            End If                               ' ColumnFilter <> "Empty"
        End If                                   ' ColumnDesignator <> "Empty"
    Else                                         ' Empty RowDesignator
        If ColumnDesignator <> "Empty" Then      ' Valid ColumnDesignator
            If ColumnFilter <> "Empty" Then      ' Valid ColumnFilter
                Select Case RowCount
                Case 0                           ' unspecified row, 1 column, filter=0 rows; no data
                    GetData.Valid = "Empty"
                Case 1                           ' unspecified row, 1 column, filter=1 row; one cell
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
    
    GetData.Headers = ThisSearchTable.Headers

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

    ' Purpose
    '   Determine if RowDesignator is valid
    '   Return the valid row number or "Empty" if valid
    '   Return ValidRowDesignator = "Error RowDesignator Out of Range" or
    '       "Error Row Designator" if invalid
    ' Assumptions
    
    Const Routine_Name As String = Module_Name & "ValidRowDesignator"
    On Error GoTo ErrorHandler
    
    ' Verify that SearchTable is valid
    If Left$(SearchTable.Valid, 5) = "Error" Then
        ValidRowDesignator.Valid = "Error Table"
        Exit Function
    End If
    
    If RowDesignator = "Empty" Then
        ' Empty is a valid entry; means return an entire column
        ValidRowDesignator = "Empty"
    Else
        Select Case VarType(RowDesignator)
        Case vbInteger, vbLong
            ' RowDesignator is numeric
            ' Verify that RowDesignator is in the range of the table's rows
            If RowDesignator >= 1 And RowDesignator <= UBound(SearchTable.Body, 1) Then
                ValidRowDesignator = RowDesignator
            Else
                ' RowDesignator is out of range
                ValidRowDesignator = "Error RowDesignator Out of Range"
            End If
        Case Else
            ' Erroneous RowDesignator data type
            ValidRowDesignator = "Error Row Designator"
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
    
    ' Purpose
    '   Determine if ColumnDesignator or "Empty" if valid
    '   Return the valid column name and number if valid
    '   Return ValidColumnDesignator.ColumnName = "Error Column Name Not Found" if invalid
    ' Assumptions
    
    Const Routine_Name As String = Module_Name & "ValidColumnDesignator"
    On Error GoTo ErrorHandler
    
    ' Verify that SearchTable is valid
    If Left$(SearchTable.Valid, 5) = "Error" Then
        ValidColumnDesignator.ColumnName = "Error Table"
        Exit Function
    End If
    
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
        
    ' Purpose
    '   Determines if the ColumnFilter is valid
    '   If ColumnFilter valid, returns the filtered array
    '   If ColumnFilter is invalid, returns ValidFilter.Valid to an error message
    ' Assumptions
    
    Const Routine_Name As String = Module_Name & "ValidFilter"
    On Error GoTo ErrorHandler
    
    ' Verify that SearchTable is valid
    If Left$(SearchTable.Valid, 5) = "Error" Then
        ValidFilter.Valid = "Error Table"
        Exit Function
    End If
    
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
            ' Iterate through ColumnFilter looking for "=", "<", or ">"
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
                    ValidFilter.Valid = "Error filter"
                    Exit Function
                End If
                ColumnNumber = TempValidColumnDesignator.ColumnNumber
                
                If NextChar = " " Then
                    Operand = Mid$(ColumnFilter, I + 2, Len(ColumnFilter) - I - 1)
                Else
                    Operand = Mid$(ColumnFilter, I + 1, Len(ColumnFilter) - I)
                End If
                
            Case "<"                             ' ThisChar = "<"
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
                        ValidFilter.Valid = "Error filter"
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
                        ValidFilter.Valid = "Error filter"
                        Exit Function
                    End If
                    ColumnNumber = TempValidColumnDesignator.ColumnNumber
                
                    Operand = Mid$(ColumnFilter, StartOfOperand, Len(ColumnFilter) - StartOfOperand + 1)
                    Exit For
                Case Else                        ' NextChar <> "=" and NextChar <> ">"
                    Operator = "<"
                    
                    StartOfOperand = I + 1
                    While Mid$(ColumnFilter, StartOfOperand, 1) = " "
                        StartOfOperand = StartOfOperand + 1
                    Wend
                    
                    ColumnName = Mid$(ColumnFilter, 1, EndOfColumnName)
                
                    TempValidColumnDesignator = ValidColumnDesignator(SearchTable, ColumnName)
                    If TempValidColumnDesignator.ColumnName = "Error" Then
                        Exit Function
                    End If
                    ColumnNumber = TempValidColumnDesignator.ColumnNumber
                
                    Operand = Mid$(ColumnFilter, StartOfOperand, Len(ColumnFilter) - StartOfOperand + 1)
                    Exit For
                End Select
            Case ">"                             ' ThisChar = ">"
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
                        ValidFilter.Valid = "Error filter"
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
                        ValidFilter.Valid = "Error filter"
                        Exit Function
                    End If
                    ColumnNumber = TempValidColumnDesignator.ColumnNumber
                
                    Operand = Mid$(ColumnFilter, StartOfOperand, Len(ColumnFilter) - StartOfOperand + 1)
                    Exit For
                End If
            End Select
        Next I
        
        Dim FilterCriteria As String
        FilterCriteria = " " & Operator & " " & """" & Operand & """"
        ValidFilter = SetUpSearchTable(SearchTable, FilterCriteria, ColumnNumber)
        If Left$(ValidFilter.Valid, 5) = "Error" Then
            Exit Function
        End If
    End If                                       ' ColumnFilter = "Empty"
    
    '@Ignore LineLabelNotUsed
Done:
    Exit Function
ErrorHandler:
    RaiseError Err.Number, Err.Source, Routine_Name, Err.Description

End Function

Private Function SetUpSearchTable( _
        SearchTable As TableType, _
        ByVal FilterCriteria As String, _
        ByVal FilterColumnNumber As Long _
        ) As TableType
        
    ' Purpose
    '   Filters Searchtable according to the FilterCriteria
    '   If ColumnFilter valid, returns the filtered array
    '   If ColumnFilter is invalid, sets SetUpSearchTable.Valid to an error message
    ' Assumptions
    
    Const Routine_Name As String = Module_Name & "SetUpSearchTableColl"
    On Error GoTo ErrorHandler
    
    ' Verify that SearchTable is valid
    If Left$(SearchTable.Valid, 5) = "Error" Then
        SetUpSearchTable.Valid = "Error Table"
        Exit Function
    End If
    
    Dim I As Long
    Dim SearchElement As SearchClass
    Dim Expression As String
    
    Dim SearchCollection As Collection
    Set SearchCollection = New Collection
    
    ' Put all the rows that match FilterCriteris into a collection
    For I = 1 To UBound(SearchTable.Body, 1)
        Expression = Chr$(34) & SearchTable.Body(I, FilterColumnNumber) & Chr$(34) & FilterCriteria
        If Evaluate(Expression) Then
            Set SearchElement = New SearchClass
            SearchElement.SetArray SearchTable, I
            SearchCollection.Add SearchElement
        End If
    Next I
    
    If SearchCollection.Count = 0 Then
        SetUpSearchTable.Valid = "Error No data found"
        Exit Function
    End If
    
    Dim DataArray As TableType
    ReDim DataArray.Body(SearchCollection.Count, UBound(SearchTable.Headers, 2))
    ReDim DataArray.Headers(1, UBound(SearchTable.Headers, 2))
    DataArray.Valid = "Valid"
    
    DataArray.Headers = SearchTable.Headers
    
    Dim RowArray As Variant
    ReDim RowArray(UBound(SearchTable.Headers, 2))
    
    Dim J As Long
    
    ' Extract all the elements in the collection and put them into DataArray
    I = 1
    Dim ValidRow As Variant
    For Each ValidRow In SearchCollection
        RowArray = ValidRow.GetArray
        For J = 1 To UBound(SearchTable.Headers, 2)
            DataArray.Body(I, J) = RowArray(J)
        Next J
        I = I + 1
    Next ValidRow
    
    SetUpSearchTable = DataArray
        
    '@Ignore LineLabelNotUsed
Done:
    Exit Function
ErrorHandler:
    RaiseError Err.Number, Err.Source, Routine_Name, Err.Description

End Function

Private Function GetRow( _
        SearchTable As TableType, _
        ByVal RowNum As Long _
        ) As Variant
        
    ' Purpose
    '   Returns the row designated by RowNum
    ' Assumptions
    
    Const Routine_Name As String = Module_Name & "GetRow"
    On Error GoTo ErrorHandler
    
    ' Verify that SearchTable is valid
    If Left$(SearchTable.Valid, 5) = "Error" Then
        GetRow.Valid = "Error Table"
        Exit Function
    End If
    
    Dim Ary As Variant
    ReDim Ary(1, UBound(SearchTable.Headers, 2))
    
    Dim I As Long
    For I = 1 To UBound(SearchTable.Headers, 2)
        Ary(1, I) = SearchTable.Body(RowNum, I)
    Next I
    
    GetRow = Ary
    
    '@Ignore LineLabelNotUsed
Done:
    Exit Function
ErrorHandler:
    RaiseError Err.Number, Err.Source, Routine_Name, Err.Description

End Function

Private Function GetColumn( _
        SearchTable As TableType, _
        ByVal ColumnNum As Long _
        ) As Variant
        
    ' Purpose
    '   Returns the column designated by ColumnNum
    ' Assumptions
    
    Const Routine_Name As String = Module_Name & "GetColumn"
    On Error GoTo ErrorHandler
    
    ' Verify that SearchTable is valid
    If Left$(SearchTable.Valid, 5) = "Error" Then
        GetColumn.Valid = "Error Table"
        Exit Function
    End If
    
    Dim Ary As Variant
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


