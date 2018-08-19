Attribute VB_Name = "SharedRoutines"
Option Explicit

Private Const Module_Name = "SharedRoutines."

Public Sub RaiseError( _
       ByVal ErrorNo As Long, _
       ByVal Src As String, _
       ByVal Proc As String, _
       ByVal Desc As String)

    ' https://excelmacromastery.com/vba-error-handling/
    ' Reraises an error and adds line number and current procedure name

    Dim SourceOfError As String

    ' Check if procedure where error occurs the line no and proc details
    If Src = ThisWorkbook.VBProject.Name Then
        ' Add error line number if present
        If Erl <> 0 Then
            SourceOfError = vbCrLf & "Line no: " & Erl & " "
        End If

        ' Add procedure to source
        SourceOfError = SourceOfError & vbCrLf & Proc

    Else
        ' If error has already been raised then just add on procedure name
        SourceOfError = Src & vbCrLf & Proc
    End If

    ' If the code stops here,
    ' make sure DisplayError is placed in the top most Sub
    Err.Raise ErrorNo, SourceOfError, Desc

End Sub                                          ' RaiseError

Public Sub DisplayError(ByVal Procname As String)

    ' https://excelmacromastery.com/vba-error-handling/
    ' Displays the error when it reaches the topmost sub
    ' Note: You can add a call to logging from this sub

    Dim Msg As String
    Msg = "The following error occurred: " & vbCrLf & Err.Description _
        & vbCrLf & vbCrLf & "Error Location is: "

    Msg = Msg + Err.Source & vbCrLf & Procname

    MsgBox Msg, Title:="Error"
End Sub                                          ' DisplayError

Public Function NumberOfArrayDimensions(ByVal Arr As Variant) As Long
    ' http://www.cpearson.com/excel/vbaarrays.htm
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' NumberOfArrayDimensions
    ' This function returns the number of dimensions of an array. An unallocated dynamic array
    ' has 0 dimensions. This condition can also be tested with IsArrayEmpty.
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim Ndx As Long
    Dim Res As Long

    Const RoutineName As String = Module_Name & "NumberOfArrayDimensions"
    On Error GoTo ErrorHandler

    On Error Resume Next
    Res = UBound(Arr, 2)                         ' If Arr has only one element, this will fail
    If Err.Number <> 0 Then
        NumberOfArrayDimensions = 1
        On Error GoTo ErrorHandler
        Exit Function
    End If

    On Error Resume Next
    ' Loop, increasing the dimension index Ndx, until an error occurs.
    ' An error will occur when Ndx exceeds the number of dimension
    ' in the array. Return Ndx - 1.
    Do
        Ndx = Ndx + 1
        Res = UBound(Arr, Ndx)
    Loop Until Err.Number <> 0

    On Error GoTo ErrorHandler

    NumberOfArrayDimensions = Ndx - 1

    '@Ignore LineLabelNotUsed
Done:
    Exit Function
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description

End Function                                     ' NumberOfArrayDimensions

Public Function IsArrayAllocated(ByVal Arr As Variant) As Boolean
' http://www.cpearson.com/excel/isarrayallocated.aspx
    On Error Resume Next
    IsArrayAllocated = _
                     IsArray(Arr) And _
                     Not IsError(LBound(Arr, 1)) And _
                     LBound(Arr, 1) <= UBound(Arr, 1)
    On Error GoTo 0
End Function




'Public Sub Log(ParamArray Msg() As Variant)
'    ' http://analystcave.com/vba-proper-vba-error-handling/
'    ' https://excelmacromastery.com/vba-error-handling/
'
'    Dim FileName As String
'    FileName = GetMainWorkbook.Path & "\error_log.txt"
'    Dim MsgString As Variant
'    Dim I As Long
'
'    Exit Sub
'
'    ' Archive file at certain size
'    If FileLen(FileName) > 20000 Then
'        FileCopy FileName, _
'                 Replace(FileName, ".txt", _
'                         Format$(Now, "ddmmyyyy hhmmss.txt"))
'        Kill FileName
'    End If
'
'    ' Open the file to write
'    Dim filenumber As Long
'    filenumber = FreeFile
'    Open FileName For Append As #filenumber
'
'    MsgString = Msg(LBound(Msg))
'    For I = LBound(Msg) + 1 To UBound(Msg)
'        MsgString = "," & MsgString & Msg(I)
'    Next I
'
'    Print #filenumber, Now & ":  " & MsgString
'
'    Close #filenumber
'
'End Sub                                          ' Log

