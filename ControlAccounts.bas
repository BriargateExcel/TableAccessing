Attribute VB_Name = "ControlAccounts"
Option Explicit

Private Const Module_Name As String = "ControlAccounts."

Private ControlAccountHeaders As Variant
Private ControlAccountData As Variant
Private ControlAccountTable As ListObject
Private ControlAccounts As DataAccessedType
Private ControlAccountNames As DataAccessedType

Public Sub ControlAccountsInitialize()

    Const Routine_Name As String = Module_Name & "ControlAccountsInitialize"
    On Error GoTo ErrorHandler
    
If Not IsArrayAllocated(ControlAccountData) Then
    Set ControlAccountTable = ControlAccountsSheet.ListObjects("ControlAccountTable")
    ControlAccountHeaders = ControlAccountTable.HeaderRowRange
    ControlAccountData = ControlAccountTable.DataBodyRange
    ControlAccounts = GetData(ControlAccountTable, , "Control Account")
    ControlAccountNames = GetData(ControlAccountTable, , "Control Account Name")
Stop
End If
Stop
'Erase ControlAccountData

    '@Ignore LineLabelNotUsed
Done:
    Exit Sub
ErrorHandler:
    DisplayError Routine_Name

End Sub

