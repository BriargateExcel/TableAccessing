Attribute VB_Name = "WorksheetRoutines"
Option Explicit
'@Folder("TableAccess.Worksheets.WorksheetRoutines")

Private Const Module_Name As String = "WorksheetRoutines."

Private pAllShts As WorksheetsClass

Private Function ModuleList() As Variant
    ModuleList = Array("XLAM_Module.")
End Function ' ModuleList

Public Sub WorksheetAdd( _
    ByVal WS As Variant, _
    ByVal ModuleName As String)
    
    Const RoutineName As String = Module_Name & "WorksheetAdd"

    Debug.Assert InScope(ModuleList, ModuleName, RoutineName)
    pAllShts.Add WS
End Sub ' WorksheetAdd

Public Sub WorksheetSetNewClass(ByVal ModuleName As String)
    Const RoutineName As String = Module_Name & "WorksheetSetNewClass"
    Debug.Assert InScope(ModuleList, ModuleName, RoutineName)
    Set pAllShts = New WorksheetsClass
End Sub ' WorksheetSetNewClass

