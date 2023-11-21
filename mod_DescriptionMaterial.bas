' Module: mod_DescriptionMaterial
Public Sub UpdateMaterialDescription(ByRef materialField As Control, ByVal newMaterialValue As String, Optional ByVal requeryControl As String = "")
    ' Handle null value and trim the new value
    newMaterialValue = Trim(Nz(newMaterialValue, ""))

    Dim currentValue As String
    ' Handle null value for the current field value
    currentValue = Nz(Trim(materialField.Value), "")

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

    ' Check if requery is needed
    If requeryControl <> "" Then
        ' Assuming materialField is on a subform inside the tab control on the main form
        Dim mainForm As Form
        Set mainForm = Forms("nav_LandingPage")

        Dim subForm As Form
        Set subForm = mainForm!NavigationSubform.Form

        ' Check if the control to requery is on the subform
        If Not subForm.Controls(requeryControl) Is Nothing Then
            subForm.Controls(requeryControl).Requery
        Else
            MsgBox "Control to requery not found on the subform.", vbExclamation, "Requery Error"
        End If
    End If
End Sub

Public Sub SetControlVisibility(controlName As String, isVisible As Boolean)
    Dim visibilityValue As String
    visibilityValue = IIf(isVisible, "-1", "0") ' "-1" for visible, "0" for not visible
    DoCmd.SetProperty controlName, acPropertyVisible, visibilityValue
End Sub

Public Sub SetControlVisibility(controlName As String, isVisible As Boolean)
    Dim visibilityValue As String
    visibilityValue = IIf(isVisible, "-1", "0") ' "-1" for visible, "0" for not visible
    DoCmd.SetProperty controlName, acPropertyVisible, visibilityValue
End Sub
