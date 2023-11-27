' Module: mod_DescriptionFeatures
Public Sub UpdateFeaturesDescription(ByRef featureField As Control, ByVal newFeatureValue As String, Optional ByVal requeryControl As String = "")
    ' Trim the new value and the existing field value
    newFeatureValue = Trim(newFeatureValue)
    Dim currentValue As String
    currentValue = Trim(featureField.Value)

    ' Update the field value based on the specified conditions
    If Left(currentValue, 2) = "; " Then
        ' If the field starts with " ; ", remove it and then append the new value
        featureField.Value = Mid(currentValue, 3) & "; " & newFeatureValue
    ElseIf Left(currentValue, 1) = " " Then
        ' If the field starts with " ", remove it and then append the new value
        featureField.Value = Mid(currentValue, 2) & "; " & newFeatureValue
    ElseIf currentValue <> "" Then
        ' If the field contains text, append the new value with a prefix of "; "
        featureField.Value = currentValue & "; " & newFeatureValue
    Else
        ' If the field is empty, just add the new value
        featureField.Value = newFeatureValue
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
