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

Public Function IsArrayAllocated(ByVal Arr As Variant) As Boolean
    ' http://www.cpearson.com/excel/isarrayallocated.aspx
    On Error Resume Next
    IsArrayAllocated = _
                     IsArray(Arr) And _
                     Not IsError(LBound(Arr, 1)) And _
                     LBound(Arr, 1) <= UBound(Arr, 1)
    On Error GoTo 0
End Function


