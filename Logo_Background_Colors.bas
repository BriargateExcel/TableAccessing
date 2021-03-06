Attribute VB_Name = "Logo_Background_Colors"
Option Explicit
'@Folder("TableAccess.Basics.Logo_Background_Colors")

Private Const Module_Name As String = "Logo_Background_Colors."

Private Const DarkestColor = &H763232 ' AF Dark Blue
Private Const LightestColor = &HE7E2E2 ' AF Light Gray
'Private Const DarkestColor = vbButtonFace
'Private Const LightestColor = vbButtonText

Public Sub DisableButton(ByVal Btn As MSForms.CommandButton)
    Btn.Enabled = False
End Sub ' DisableButton

Public Sub EnableButton(ByVal Btn As MSForms.CommandButton)
    Btn.Enabled = True
End Sub ' EnableButton

Public Sub HighLightButton(ByVal Btn As MSForms.CommandButton)
    Btn.ForeColor = DarkestColor
    Btn.BackColor = LightestColor
    Btn.Enabled = True
End Sub ' HighLightButton

Public Sub HighLightControl(ByVal Ctl As Control)
    Ctl.ForeColor = DarkestColor
    Ctl.BackColor = LightestColor
End Sub ' HighLightControl

Public Sub LowLightButton(ByVal Btn As MSForms.CommandButton)
    Btn.ForeColor = LightestColor
    Btn.BackColor = DarkestColor
    Btn.Enabled = True
End Sub ' LowLightButton

Public Sub LowLightControl(ByVal Ctl As Control)
    Ctl.ForeColor = LightestColor
    Ctl.BackColor = DarkestColor
End Sub ' LowLightControl

Public Sub Texture(ByRef Tbl As TableClass)
    Const RoutineName As String = Module_Name & "Texture"
    On Error GoTo ErrorHandler
    
    If Dir(MainWorkbook.Path & "\texture.jpg") <> vbNullString Then
        Set Tbl.Form.FormObj.Picture = LoadPicture(MainWorkbook.Path & "\texture.jpg")
    End If
'@Ignore LineLabelNotUsed
Done:
    Exit Sub
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description

End Sub

Public Function Logo( _
    ByRef Tbl As TableClass, _
    ByRef LogoHeight As Single, _
    ByRef LogoWidth As Single _
    ) As Control
    
    Dim LogoImage As Control
    
    Const RoutineName As String = Module_Name & "Logo"
    On Error GoTo ErrorHandler
    
    If Dir(MainWorkbook.Path & "\logo.jpg") <> vbNullString Then
        Set LogoImage = Tbl.Form.FormObj.Controls.Add("Forms.Image.1")
        Set LogoImage.Picture = LoadPicture(MainWorkbook.Path & "\logo.jpg")
        With LogoImage
            .PictureAlignment = fmPictureAlignmentTopRight
            .PictureSizeMode = fmPictureSizeModeZoom
            .BorderStyle = fmBorderStyleNone
            .BackStyle = fmBackStyleTransparent
            .AutoSize = True
            LogoHeight = .Height
            LogoWidth = .Width
        End With
        Set Logo = LogoImage
    Else
        LogoHeight = 0
        Set Logo = Nothing
    End If
'@Ignore LineLabelNotUsed
Done:
    Exit Function
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description

End Function


