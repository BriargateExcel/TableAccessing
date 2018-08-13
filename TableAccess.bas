Attribute VB_Name = "TableAccess"
Option Explicit

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

Sub test()
    '    GetData Worksheets("Sheet1").ListObjects("ControlAccountTable"), 3, , "Control Account = 8J6GM15223-02A"
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
       ) As Variant
    
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
        GetData = "Error Search Table"
        Exit Function
    End If
    
    ' Verify that RowDesignator is a valid row
    Dim RowNumber As Long
    Dim TempRowDesignator As Variant
    TempRowDesignator = ValidRowDesignator(SearchTable, RowDesignator)
    If TempRowDesignator = "Error" Then
        GetData = "Error Row Designator"
        Exit Function
    Else
        ' Valid RowDesignator
        RowNumber = TempRowDesignator
    End If
    
    ' Verify that ColumnDesignator is a valid column
    Dim ColumnName As String
    Dim ColumnNumber As Long
    Dim TempColumnDesignator As ColumnDesignatorType
    TempColumnDesignator = ValidColumnDesignator(SearchTable, ColumnDesignator)
    If TempColumnDesignator.ColumnName = "Error" Then
        ColumnName = "Error"
        ColumnNumber = 0
    Else
        ColumnName = ColumnDesignator
        ColumnNumber = TempColumnDesignator.ColumnNumber
    End If
    
    ' Verify that ColumnFilter is a valid column
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
    Else
        FilterColumnName = TempFilterValues.ColumnName
        FilterColumnNumber = TempFilterValues.ColumnNumber
        FilterOperator = TempFilterValues.Operator
        FilterOperand = TempFilterValues.Operand
    End If
    
End Function

Private Function ValidSearchTable(ByVal SearchTable As ListObject) As Boolean
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
                    ValidFilter.Operand = Mid$(ColumnFilter, StartOfOperand, Len(ColumnFilter) - StartOfOperand + 1)
                    Exit For
                Case ">"
                    ValidFilter.Operator = "<>"
                    
                    StartOfOperand = I + 2
                    While Mid$(ColumnFilter, StartOfOperand, 1) = " "
                        StartOfOperand = StartOfOperand + 1
                    Wend
                    
                    ValidFilter.ColumnName = Mid$(ColumnFilter, 1, EndOfColumnName)
                    ValidFilter.Operand = Mid$(ColumnFilter, StartOfOperand, Len(ColumnFilter) - StartOfOperand + 1)
                    Exit For
                Case Else
                    ValidFilter.Operator = "<"
                    
                    StartOfOperand = I + 1
                    While Mid$(ColumnFilter, StartOfOperand, 1) = " "
                        StartOfOperand = StartOfOperand + 1
                    Wend
                    
                    ValidFilter.ColumnName = Mid$(ColumnFilter, 1, EndOfColumnName)
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
                    ValidFilter.Operand = Mid$(ColumnFilter, StartOfOperand, Len(ColumnFilter) - StartOfOperand + 1)
                    Exit For
                Else
                    ValidFilter.Operator = ">"
                    
                    StartOfOperand = I + 1
                    While Mid$(ColumnFilter, StartOfOperand, 1) = " "
                        StartOfOperand = StartOfOperand + 1
                    Wend
                    
                    ValidFilter.ColumnName = Mid$(ColumnFilter, 1, EndOfColumnName)
                    ValidFilter.Operand = Mid$(ColumnFilter, StartOfOperand, Len(ColumnFilter) - StartOfOperand + 1)
                    Exit For
                End If
            End Select
        Next I
    End If
    
End Function


