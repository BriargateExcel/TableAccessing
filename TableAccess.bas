Attribute VB_Name = "TableAccess"
Option Explicit

Private Const Module_Name As String = "TableAccess."

Private Const TempWorkSheetName As String = "TempWorkSheet"

Public Type DataAccessedType
    RowCount As Long
    ColumnCount As Long
    Data As Variant
End Type

Private Type ColumnDesignatorType
    ColumnName As String
    ColumnNumber As Long
End Type

Private Sub test()

    Const Routine_Name As String = Module_Name & "test"
    On Error GoTo ErrorHandler
    
    GetData Sheet1.ListObjects("ControlAccountTable"), , , "CAM = Dye"
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

    '@Ignore LineLabelNotUsed
Done:
    Exit Sub
ErrorHandler:
    DisplayError Routine_Name

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
    
    ' Verify that Searchtable is valid
    If ValidSearchTable(SearchTable) Then        ' Valid Table
    Else                                         ' Invalid Table
        GetData.Data = "Error Search Table"
        Exit Function
    End If
    
    ' Verify that RowDesignator is valid
    Dim RowNumber As Long
    Dim TempRowDesignator As Variant
    TempRowDesignator = ValidRowDesignator(SearchTable, RowDesignator)
    Select Case TempRowDesignator
    Case "Error"                                 ' Invalid RowDesignator
        GetData.Data = "Error Row Designator"
        Exit Function
    Case "Empty"                                 ' Empty RowDesignator
        RowNumber = 0
    Case Else                                    ' Conclude that RowDesignator must be a number
        RowNumber = TempRowDesignator
        GetData.RowCount = 1
    End Select
    
    ' Verify that ColumnDesignator is valid
    Dim ColumnName As String
    Dim ColumnNumber As Long
    Dim TempColumnDesignator As ColumnDesignatorType
    TempColumnDesignator = ValidColumnDesignator(SearchTable, ColumnDesignator)
    Select Case TempColumnDesignator.ColumnName
    Case "Error"                                 ' Invalid RowDesignator
        GetData.Data = "Error Column Designator"
        Exit Function
    Case "Empty"                                 ' Empty RowDesignator
        ColumnNumber = 0
    Case Else                                    ' Conclude that RowDesignator must be a number
        GetData.ColumnCount = 1
        ColumnNumber = TempColumnDesignator.ColumnNumber
    End Select
    ColumnName = TempColumnDesignator.ColumnName
    
    ' Verify that ColumnFilter is valid and set up the SearchTable
    Dim ThisSearchTable As Variant
    Set ThisSearchTable = ValidFilter(SearchTable, ColumnFilter)
    If ThisSearchTable = "Error" Then
        GetData.Data = "Error Bad ColumnFilter"
        Exit Function
    Else
        Dim RowCount As Long
        RowCount = ThisSearchTable.DataBodyRange.Columns(1).Cells.Count
    End If
    
    If RowDesignator <> "Empty" Then             ' Valid RowDesignator
        If ColumnDesignator <> "Empty" Then      ' Valid ColumnDesignator
            If ColumnFilter <> "Empty" Then      ' Valid ColumnFilter
                Select Case RowCount
                Case 0                           ' 1 row, 1 column, filter=0 rows; no data
                    GetData.Data = "Empty"
                    GetData.RowCount = 0
                    GetData.ColumnCount = 0
                Case 1                           ' 1 row, 1 column, filter=1 row; one cell
                    GetData.Data = ThisSearchTable.SearchTable.DataBodyRange(RowNumber, ColumnNumber)
                Case Else                        ' 1 row, 1 column, filter=multiple rows; one cell
                    GetData.Data = ThisSearchTable.SearchTable.DataBodyRange(RowNumber, ColumnNumber)
                End Select
            Else                                 ' 1 row, 1 column, empty filter; one cell
                GetData.Data = ThisSearchTable.SearchTable.DataBodyRange(RowNumber, ColumnNumber)
            End If                               ' ColumnFilter <> "Empty"
        Else                                     ' Empty ColumnDesignator
            GetData.ColumnCount = ThisSearchTable.SearchTable.DataBodyRange.Columns.Count
            If ColumnFilter <> "Empty" Then
                ' Valid ColumnFilter
                Select Case RowCount
                Case 0                           ' 1 row, unspecified columns, filter=0 rows; no data
                    GetData.Data = "Empty"
                Case 1                           ' 1 row, unspecified columns, filter=1 row; one row
                    GetData.Data = ThisSearchTable.SearchTable.DataBodyRange.Rows(RowNumber)
                Case Else                        ' 1 row, unspecified columns, filter=multiple rows; one row
                    GetData.Data = ThisSearchTable.SearchTable.DataBodyRange.Rows(RowNumber)
                End Select
            Else                                 ' 1 row, unspecified column, empty filter; entire row
                GetData.Data = ThisSearchTable.SearchTable.DataBodyRange.Rows(RowNumber)
            End If                               ' ColumnFilter <> "Empty"
        End If                                   ' ColumnDesignator <> "Empty"
    Else                                         ' Empty RowDesignator
        If ColumnDesignator <> "Empty" Then      ' Valid ColumnDesignator
            If ColumnFilter <> "Empty" Then      ' Valid ColumnFilter
                Select Case RowCount
                Case 0                           ' unspecified row, 1 column, filter=0 rows; no data
                    GetData.Data = "Empty"
                    GetData.RowCount = 0
                    GetData.ColumnCount = 0
                Case 1                           ' unspecified row, 1 column, filter=1 row, one cell
                    GetData.Data = ThisSearchTable.DataBodyRange.Columns(ColumnNumber)
                Case Else                        ' unspecified row, 1 column, filter=multiple rows; one column
                    GetData.Data = ThisSearchTable.DataBodyRange.Columns(ColumnNumber)
                    GetData.RowCount = RowCount
                End Select
            Else                                 ' empty row, one column, empty filter; one entire column
                GetData.Data = ThisSearchTable.DataBodyRange.Columns(ColumnNumber)
            End If                               ' ColumnFilter <> "Empty"
        Else                                     ' Empty ColumnDesignator
            GetData.ColumnCount = ThisSearchTable.DataBodyRange.Columns.Count
            If ColumnFilter <> "Empty" Then      ' Valid ColumnFilter
                Select Case RowCount
                Case 0                           ' unspecified row, unspecified column, filter=0 rows; no data
                    GetData.Data = "Empty"
                    GetData.RowCount = 0
                    GetData.ColumnCount = 0
                Case 1                           ' unspecified row, unspecified column, filter=1 row; one row, all columns
                    GetData.Data = ThisSearchTable.DataBodyRange
                Case Else                        ' unspecified row, unspecified column, filter=multiple rows; multiple rows, all columns
                    GetData.Data = ThisSearchTable.DataBodyRange
                    GetData.RowCount = RowCount
                End Select
            Else                                 ' empty row, empty column, empty filter; entire table
                GetData.Data = ThisSearchTable.DataBodyRange
            End If                               ' ColumnFilter <> "Empty"
        End If                                   ' ColumnDesignator <> "Empty"
    End If                                       ' RowDesignator <> "Empty"
    
    
    ' Delete the temporary worksheet
    Dim TempDisplayAlerts As Boolean
    TempDisplayAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False
    Workbooks(Worksheets(TempWorkSheetName).Parent.Name).Worksheets(TempWorkSheetName).Delete
    Application.DisplayAlerts = TempDisplayAlerts
    
    '@Ignore LineLabelNotUsed
Done:
    Exit Function
ErrorHandler:
    RaiseError Err.Number, Err.Source, Routine_Name, Err.Description

End Function

Private Function ValidSearchTable(ByVal SearchTable As ListObject) As Boolean
    ' Assumes SearchTable is a valid ListObject
    
    Const Routine_Name As String = Module_Name & "ValidSearchTable"
    On Error GoTo ErrorHandler
    
    Dim CheckForValidSearchTable As String
    ValidSearchTable = True
    On Error GoTo 0
    CheckForValidSearchTable = SearchTable.Name
    If Err.Number <> 0 Then
        ValidSearchTable = False
        Exit Function
    End If
    On Error GoTo 0
    
    '@Ignore LineLabelNotUsed
Done:
    Exit Function
ErrorHandler:
    RaiseError Err.Number, Err.Source, Routine_Name, Err.Description

End Function

Private Function ValidRowDesignator( _
        ByVal SearchTable As ListObject, _
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
    
    '@Ignore LineLabelNotUsed
Done:
    Exit Function
ErrorHandler:
    RaiseError Err.Number, Err.Source, Routine_Name, Err.Description

End Function

Private Function ValidColumnDesignator( _
        ByVal SearchTable As ListObject, _
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
    
    '@Ignore LineLabelNotUsed
Done:
    Exit Function
ErrorHandler:
    RaiseError Err.Number, Err.Source, Routine_Name, Err.Description

End Function

Private Function ValidFilter( _
        ByVal SearchTable As ListObject, _
        ByVal ColumnFilter As String _
        ) As Variant
    
    ' Assumes SearchTable is a valid ListObject

    Const Routine_Name As String = Module_Name & "ValidFilter"
    On Error GoTo ErrorHandler
    
    Set ValidFilter = SearchTable
        
    If ColumnFilter = "Empty" Then
        ' "Empty" is a valid ColumnFilter value
        ValidFilter = "Empty"
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
                    ValidFilter = "Error Bad Column Filter"
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
                        ValidFilter = "Error Bad Column Filter"
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
                        ValidFilter = "Error Bad Column Filter"
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
                        ValidFilter = "Error Bad Column Filter"
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
                        ValidFilter = "Error Bad Column Filter"
                        Exit Function
                    End If
                    ColumnNumber = TempValidColumnDesignator.ColumnNumber
                
                    Operand = Mid$(ColumnFilter, StartOfOperand, Len(ColumnFilter) - StartOfOperand + 1)
                    Exit For
                End If
            End Select
        Next I
    End If
    
    Dim FilterCriteria As String
    FilterCriteria = Operator & Operand
    Set ValidFilter = GetFilteredData(FilterCriteria, ColumnNumber, SearchTable)
    
    '@Ignore LineLabelNotUsed
Done:
    Exit Function
ErrorHandler:
    RaiseError Err.Number, Err.Source, Routine_Name, Err.Description

End Function

Private Function GetFilteredData( _
    ByVal FilterCriteria As String, _
    ByVal FilterColumnNumber As Long, _
    ByVal SearchTable As ListObject _
    ) As ListObject
    
    Const Routine_Name As String = Module_Name & "GetFilteredData"
    On Error GoTo ErrorHandler
    
        Dim CurrentWorkBook As Workbook
        Set CurrentWorkBook = Workbooks(Worksheets(SearchTable.Parent.Name).Parent.Name)
        
        Dim CurrentSheet As Worksheet
        Set CurrentSheet = CurrentWorkBook.ActiveSheet
        Worksheets(SearchTable.Parent.Name).Activate
        
        SearchTable.DataBodyRange(1, 1).Activate
        On Error Resume Next                     ' ShowAllData throws an error if the table is already completely unfiltered
        ThisWorkbook.Worksheets(SearchTable.Parent.Name).ShowAllData
        On Error GoTo ErrorHandler
        
        On Error Resume Next
        CurrentWorkBook.Worksheets(TempWorkSheetName).Activate
        If Err.Number = 0 Then
            Dim TempDisplayAlerts As Boolean
            TempDisplayAlerts = Application.DisplayAlerts
            Application.DisplayAlerts = False
            CurrentWorkBook.Worksheets(TempWorkSheetName).Delete
            Application.DisplayAlerts = TempDisplayAlerts
        End If
        On Error GoTo ErrorHandler
        
        CurrentWorkBook.Worksheets(SearchTable.Parent.Name).Activate
        SearchTable.Range.AutoFilter Field:=FilterColumnNumber, Criteria1:=FilterCriteria
        SearchTable.DataBodyRange.SpecialCells(xlCellTypeVisible).Copy
        
        Dim TempWorkSheet As Worksheet
        Set TempWorkSheet = CurrentWorkBook.Sheets.Add(After:=CurrentWorkBook.Worksheets(Worksheets.Count))
        TempWorkSheet.Name = TempWorkSheetName
        
        TempWorkSheet.Range("$A$1").PasteSpecial
        TempWorkSheet.ListObjects.Add(xlSrcRange, TempWorkSheet.UsedRange, , xlYes).Name = "TempTable"
        Set GetFilteredData = TempWorkSheet.ListObjects("TempTable")
        
        CurrentWorkBook.Worksheets(SearchTable.Parent.Name).Activate
        SearchTable.DataBodyRange(1, 1).Activate
        On Error Resume Next                     ' ShowAllData throws an error if the table is already completely unfiltered
        CurrentWorkBook.Worksheets(SearchTable.Parent.Name).ShowAllData
        On Error GoTo ErrorHandler
        
        CurrentSheet.Activate
        
    '@Ignore LineLabelNotUsed
Done:
    Exit Function
ErrorHandler:
    RaiseError Err.Number, Err.Source, Routine_Name, Err.Description

End Function
