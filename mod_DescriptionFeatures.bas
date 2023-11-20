' Module: mod_DescriptionFeatures
Public Sub UpdateFeaturesDescription(ByRef featureField As Control, ByVal newFeatureValue As String, Optional ByVal requeryControl As String = "")
    ' Trim the new value and the existing field value
    newMaterialValue = Trim(newMaterialValue)
    Dim currentValue As String
    currentValue = Trim(materialField.Value)

    ' Update the field value based on the specified conditions
    If Left(currentValue, 2) = "; " Then
        ' If the field starts with " ; ", remove it and then append the new value
        materialField.Value = Mid(currentValue, 3) & "; " & newMaterialValue
    ElseIf Left(currentValue, 1) = " " Then
        ' If the field starts with " ", remove it and then append the new value
        materialField.Value = Mid(currentValue, 2) & "; " & newMaterialValue
    ElseIf currentValue <> "" Then
        ' If the field contains text, append the new value with a prefix of "; "
        materialField.Value = currentValue & "; " & newMaterialValue
    Else
        ' If the field is empty, just add the new value
        materialField.Value = newMaterialValue
    End If

    ' Requery if needed
    If requeryControl <> "" Then
        Dim parentForm As Form
        Set parentForm = featureField.Form

        ' Requery the control on the form
        If Not parentForm.Controls(requeryControl) Is Nothing Then
            parentForm.Controls(requeryControl).Requery
        End If
    End If
End Sub

Public Sub SetFeatureControlVisibility(controlName As String, isVisible As Boolean)
    Dim visibilityValue As String
    visibilityValue = IIf(isVisible, "-1", "0") ' "-1" for visible, "0" for not visible
    DoCmd.SetProperty controlName, acPropertyVisible, visibilityValue
End Sub
