VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EventsClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'@Folder("TableAccess.Events.EventsClass")

Private Const Module_Name As String = "EventsClass."

Private pAllEvents As Scripting.Dictionary

Private Sub Class_Initialize()
    Debug.Assert Initializing

    Set pAllEvents = New Scripting.Dictionary
End Sub ' Class_Initialize

Public Property Get Count() As Long: Count = pAllEvents.Count: End Property

Public Property Get Item(ByVal vIndex As Variant) As EventClass
'Attribute Item.VB_UserMemId = 0
    
    Const RoutineName As String = Module_Name & "Get Item"
    
    On Error Resume Next
    Set Item = pAllEvents.Items()(vIndex)
    If Err.Number = 0 Then
        On Error GoTo 0
        Exit Property
    End If
    
    On Error Resume Next
    Set Item = pAllEvents(vIndex)
    On Error GoTo 0

'@Ignore LineLabelNotUsed
Done:
    Exit Property
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description

End Property

Public Sub Add(ByVal Evt As EventClass)

    Debug.Assert Initializing

    Const RoutineName As String = Module_Name & "Add"
    On Error GoTo ErrorHandler
           
    If Evt.Name <> vbNullString Then pAllEvents.Add Evt.Name, Evt
        
'@Ignore LineLabelNotUsed
Done:
    Exit Sub
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
    
End Sub ' Add

Public Sub Remove(ByVal vIndex As Variant)

    Debug.Assert Initializing
    
    Const RoutineName As String = Module_Name & "Remove"
    On Error GoTo ErrorHandler
   
    If CStr(vIndex) = "*" Then
        Set pAllEvents = Nothing
        Set pAllEvents = New Collection
    Else
        If Not pAllEvents.Exists(vIndex) Then Err.Raise 9
        pAllEvents.Remove vIndex
    End If

'@Ignore LineLabelNotUsed
Done:
    Exit Sub
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
    
End Sub ' Remove
