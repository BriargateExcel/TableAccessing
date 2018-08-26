Attribute VB_Name = "ControlAccounts"
Option Explicit

Private Const Module_Name As String = "ControlAccounts."

Private ControlAccountHeaders As Variant
Private ControlAccountData As TableType
Private ControlAccounts As TableType
Private ControlAccountNames As TableType

Public Sub ControlAccountsInitialize()

    Const Routine_Name As String = Module_Name & "ControlAccountsInitialize"
    On Error GoTo ErrorHandler

If Not IsArrayAllocated(ControlAccountData.Body) Then
    ControlAccountData.Headers = ControlAccountsSheet.ListObjects("ControlAccountTable").HeaderRowRange
    ControlAccountData.Body = ControlAccountsSheet.ListObjects("ControlAccountTable").DataBodyRange
    ControlAccountData.Valid = "Valid"
    ControlAccounts = GetData(ControlAccountData, , "Control Account")
    ControlAccountNames = GetData(ControlAccountData, , "Control Account Name")
End If
'Erase ControlAccountData.Body

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
    
    Dim Temp As TableType
    
    Temp = GetData(ControlAccountData, , , "CAM = Dye")
    Temp = GetData(ControlAccountData, 3, , "Control Account=8J6GM15223-02A")
    Temp = GetData(ControlAccountData, 3, , "Control Account =8J6GM15223-02A")
    Temp = GetData(ControlAccountData, 3, , "Control Account= 8J6GM15223-02A")
    
    Temp = GetData(ControlAccountData, 3, , "Control Account < 8J6GM15223-02A")
    Temp = GetData(ControlAccountData, 3, , "Control Account<8J6GM15223-02A")
    Temp = GetData(ControlAccountData, 3, , "Control Account <= 8J6GM15223-02A")
    Temp = GetData(ControlAccountData, 3, , "Control Account <> 8J6GM15223-02A")
    Temp = GetData(ControlAccountData, 3, , "Control Account<=8J6GM15223-02A")
    Temp = GetData(ControlAccountData, 3, , "Control Account<>8J6GM15223-02A")
    Temp = GetData(ControlAccountData, 3, , "Control Account <8J6GM15223-02A")
    Temp = GetData(ControlAccountData, 3, , "Control Account <=8J6GM15223-02A")
    Temp = GetData(ControlAccountData, 3, , "Control Account <>8J6GM15223-02A")
    Temp = GetData(ControlAccountData, 3, , "Control Account< 8J6GM15223-02A")
    Temp = GetData(ControlAccountData, 3, , "Control Account<= 8J6GM15223-02A")
    Temp = GetData(ControlAccountData, 3, , "Control Account<> 8J6GM15223-02A")
    
    Temp = GetData(ControlAccountData, 3, , "Control Account > 8J6GM15223-02A")
    Temp = GetData(ControlAccountData, 3, , "Control Account >= 8J6GM15223-02A")
    
    Temp = GetData(ControlAccountData, 3, , "Control Account>8J6GM15223-02A")
    Temp = GetData(ControlAccountData, 3, , "Control Account>=8J6GM15223-02A")
    
    Temp = GetData(ControlAccountData, 3, , "Control Account >8J6GM15223-02A")
    Temp = GetData(ControlAccountData, 3, , "Control Account >=8J6GM15223-02A")
    
    Temp = GetData(ControlAccountData, 3, , "Control Account> 8J6GM15223-02A")
    Temp = GetData(ControlAccountData, 3, , "Control Account>= 8J6GM15223-02A")

    '@Ignore LineLabelNotUsed
Done:
    Exit Sub
ErrorHandler:
    DisplayError Routine_Name

End Sub

