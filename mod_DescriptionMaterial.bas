' Module: mod_DescriptionMaterial

Public Sub UpdateMaterialDescription(ByRef materialField As String, ByVal newMaterialValue As String, Optional ByVal requeryControl As String = "")
    If IsNull(materialField) Then
        materialField = newMaterialValue
    Else
        materialField = materialField & "; " & newMaterialValue
    End If

    If requeryControl <> "" Then
        DoCmd.Requery requeryControl
    End If
End Sub

Public Sub SetControlVisibility(controlName As String, isVisible As Boolean)
    Dim visibilityValue As String
    visibilityValue = IIf(isVisible, "-1", "0") ' "-1" for visible, "0" for not visible
    DoCmd.SetProperty controlName, acPropertyVisible, visibilityValue
End Sub
