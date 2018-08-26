Attribute VB_Name = "ControlAccounts"
Option Explicit

Private Const Module_Name As String = "ControlAccounts."

Private ControlAccountTable As TableType
Private DataTable As TableType

Public Sub ControlAccountsInitialize()

    Const Routine_Name As String = Module_Name & "ControlAccountsInitialize"
    On Error GoTo ErrorHandler

    If Not IsArrayAllocated(ControlAccountTable.Body) Then
        ControlAccountTable.Headers = ControlAccountsSheet.ListObjects("ControlAccountTable").HeaderRowRange
        ControlAccountTable.Body = ControlAccountsSheet.ListObjects("ControlAccountTable").DataBodyRange
        ControlAccountTable.Valid = "Valid"
    End If

    '@Ignore LineLabelNotUsed
Done:
    Exit Sub
ErrorHandler:
    DisplayError Routine_Name

End Sub

Public Sub DataTableInitialize()

    Const Routine_Name As String = Module_Name & "DataTableInitialize"
    On Error GoTo ErrorHandler

    If Not IsArrayAllocated(DataTable.Body) Then
        DataTable.Headers = DataSheet.ListObjects("DataTable").HeaderRowRange
        DataTable.Body = DataSheet.ListObjects("DataTable").DataBodyRange
        DataTable.Valid = "Valid"
    End If

    '@Ignore LineLabelNotUsed
Done:
    Exit Sub
ErrorHandler:
    DisplayError Routine_Name

End Sub

Private Sub test()

    Const Routine_Name As String = Module_Name & "test"
    On Error GoTo ErrorHandler
    
    ControlAccountsInitialize
    DataTableInitialize
    
    Dim Temp As TableType
    
    ' Get a specific row
    Temp = GetData(ControlAccountTable, 3)
    
    ' Get a specific column
    Temp = GetData(ControlAccountTable, , "Control Account")
    
    ' Get a specific cell
    Temp = GetData(ControlAccountTable, , , "Control Account=8G3SN04311-03")
    
    ' Get a collection of rows
    Temp = GetData(ControlAccountTable, , , "Control Account <> 8G3SN04311-03")
    Temp = GetData(ControlAccountTable, , , "Control Account < 8G3SN04311-03")
    Temp = GetData(ControlAccountTable, , , "Control Account <= 8G3SN04311-03")
    Temp = GetData(ControlAccountTable, , , "Control Account>8G3SN04311-03")
    Temp = GetData(ControlAccountTable, , , "Control Account>=8G3SN04311-03")
    
    '@Ignore LineLabelNotUsed
Done:
    Exit Sub
ErrorHandler:
    DisplayError Routine_Name

End Sub

