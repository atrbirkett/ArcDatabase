Attribute VB_Name = "UpdateDescriptionMaterial"
Option Compare Database
Public Sub UpdateDescriptionMaterial(ByRef Description_Field As String, ByVal Material_Value As String, Optional ByVal RequeryControlName As String = "")
    ' Check if the Description_Field is null and set it accordingly
    If IsNull(Description_Field) Then
        Description_Field = Material_Value
    Else
        Description_Field = Description_Field & "; " & Material_Value
    End If
    ' Requery the control if a control name is provided
    If RequeryControlName <> "" Then
        Forms(Application.CurrentForm.Name).Controls(RequeryControlName).Requery
    End If
End Sub