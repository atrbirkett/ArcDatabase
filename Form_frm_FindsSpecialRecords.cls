'
'   ----- ACCESSION & ARCHIVE SECTION ------------------------------------------------------------------------------------
'
Private Sub ID_BagIDField_AfterUpdate()
    On Error GoTo ErrorHandler
    Me.ID_AccessionNumber = "2022.HARP/" & ID_BoxID.Value & "SF" & Right("0000" & CStr(ID_FindID), 4)
    Exit Sub
ErrorHandler:
    MsgBox "An error occurred: " & Err.Description
End Sub
Private Sub AccessionID_AfterUpdate()
    Me.ID_AccessionNumber = "2022.HARP/" & ID_BoxID.Value & "SF" & Right("0000" & CStr(ID_FindID), 4)
End Sub
Private Sub AccessionID_Click()
    Me.ID_AccessionNumber = "2022.HARP/" & ID_BoxID.Value & "SF" & Right("0000" & CStr(ID_FindID), 4)
End Sub
Private Sub BoxID_AfterUpdate()
    Me.AccessionID = "2022.HARP/" & ID_BoxID.Value & "/SF" & Right("0000" & CStr(ID_FindID), 4)
End Sub
Private Sub cmdAssemblage_Click()
    DoCmd.OpenForm "frm_Assemblages", acNormal, "", "", acFormEdit, acNormal
End Sub
'
'   ----- DATING SECTION ------------------------------------------------------------------------------------
'
Private Sub Dating_Period_AfterUpdate()
    Dating_MinYear = [Dating_Period].[Column](2)
    Dating_MaxYear = [Dating_Period].[Column](3)
End Sub
'
'   ----- MATERIAL SECTION ------------------------------------------------------------------------------------
'
'   MATERIAL TYPE
'   MATERIAL TYPE LEVEL ONE
'
Private Sub cmd_Description_MaterialType_Level01_Click()
    mod_DescriptionMaterial.UpdateMaterialDescription Me.Description_MaterialType, Me.Description_MaterialType_Level01.Value, "Description_MaterialType"
End Sub
Private Sub Description_MaterialType_Level01_AfterUpdate()
    DoCmd.Requery "Description_MaterialClass_Level01"
    DoCmd.Requery "Description_MaterialClass_Level02"
    DoCmd.Requery "Description_MaterialClass_Level03"
    DoCmd.Requery "Description_MaterialType_Level01"
End Sub
Private Sub Description_MaterialType_Level01_Click()
    mod_DescriptionMaterial.SetControlVisibility "Description_MaterialClass_Level02", True
    mod_DescriptionMaterial.SetControlVisibility "cmd_Description_MaterialClass_Level02", True
    Description_MatDefinition = "Definition: " & [Description_MaterialType_Level01].[Column](1)
End Sub
'
'   MATERIAL CLASSSIFICATION
'   MATERIAL CLASSSIFICATION LEVEL ONE
'
Private Sub cmd_Description_MaterialClass_Level01_Click()
    mod_DescriptionMaterial.UpdateMaterialDescription Me.Description_MaterialClass, Me.Description_MaterialClass_Level01.Value, "Description_MaterialClass"
End Sub
Private Sub Description_MaterialClass_Level01_AfterUpdate()
    DoCmd.Requery "Description_MaterialClass_Level01"
    DoCmd.Requery "Description_MaterialClass_Level02"
    DoCmd.Requery "Description_MaterialClass_Level03"
    DoCmd.Requery "Description_MaterialType_Level01"
End Sub
Private Sub Description_MaterialClass_Level01_Click()
    mod_DescriptionMaterial.SetControlVisibility "Description_MaterialClass_Level02", True
    mod_DescriptionMaterial.SetControlVisibility "cmd_Description_MaterialClass_Level02", True
    Description_MatDefinition = "Definition: " & [Description_MaterialClass_Level01].[Column](1)
End Sub
'
'   MATERIAL CLASSSIFICATION LEVEL TWO
'
Private Sub cmd_Description_MaterialClass_Level02_Click()
    mod_DescriptionMaterial.UpdateMaterialDescription Me.Description_MaterialClass, Me.Description_MaterialClass_Level02.Value, "Description_MaterialClass"
End Sub
Private Sub Description_MaterialClass_Level02_AfterUpdate()
    DoCmd.Requery "Description_MaterialClass_Level01"
    DoCmd.Requery "Description_MaterialClass_Level02"
    DoCmd.Requery "Description_MaterialClass_Level03"
    DoCmd.Requery "Description_MaterialType_Level01"
End Sub
Private Sub Description_MaterialClass_Level02_Click()
    mod_DescriptionMaterial.SetControlVisibility "Description_MaterialClass_Level03", True
    mod_DescriptionMaterial.SetControlVisibility "cmd_Description_MaterialClass_Level03", True
    Description_MatDefinition = "Definition: " & [Description_MaterialClass_Level02].[Column](1)
End Sub
'
'   MATERIAL CLASSSIFICATION LEVEL THREE
'
Private Sub cmd_Description_MaterialClass_Level03_Click()
    mod_DescriptionMaterial.UpdateMaterialDescription Me.Description_MaterialClass, Me.Description_MaterialClass_Level03.Value, "Description_MaterialClass"
End Sub
Private Sub Description_MaterialClass_Level03_AfterUpdate()
    DoCmd.Requery "Description_MaterialClass_Level01"
    DoCmd.Requery "Description_MaterialClass_Level02"
    DoCmd.Requery "Description_MaterialClass_Level03"
    DoCmd.Requery "Description_MaterialType_Level01"
End Sub
Private Sub Description_MaterialClass_Level03_Click()
    Description_MatDefinition = "Definition: " & [Description_MaterialClass_Level03].[Column](1)
End Sub
'
'   MATERIAL CLASSSIFICATION LEVEL ALL LEVELS
'
Private Sub cmd_Description_MaterialClass_AllLevels_Click()
    mod_DescriptionMaterial.UpdateMaterialDescription Me.Description_MaterialClass, Me.Description_MaterialClass_AllLevels.Value, "Description_MaterialClass"
End Sub
Private Sub Description_MaterialClass_AllLevels_Click()
    Description_MatDefinition = "Definition: " & [Description_MaterialClass_AllLevels].[Column](1)
End Sub
'
'   MATERIAL RESET
'
Private Sub cmd_ResetMatDescription_Click()
    ' Resets fields
    Description_MaterialType = " "
    Description_MaterialType_Level01 = " "
    Description_MaterialClass = " "
    Description_MaterialClass_AllLevels = " "
    Description_MaterialClass_Level01 = " "
    Description_MaterialClass_Level02 = " "
    Description_MaterialClass_Level03 = " "
    Description_MatDefinition = " "
End Sub
'
'   ------ FEATURES SECTION ------------------------------------------------------------------------------------
'
'   FEATURES CATEGORY
'   FEATURES CATEGORY LEVEL ONE
'
Private Sub cmd_Description_FeaturesCategory_Level01_Click()
    mod_DescriptionFeatures.UpdateFeaturesDescription Me.Description_FeaturesCategory, Me.Description_FeaturesCategory_Level01.Value, "Description_FeaturesCategory"
End Sub
Private Sub Description_FeaturesCategory_Level01_AfterUpdate()
    DoCmd.Requery "Description_FeaturesType_Level01"
    DoCmd.Requery "Description_FeaturesType_Level02"
    DoCmd.Requery "Description_FeaturesType_Level03"
    DoCmd.Requery "Description_FeaturesType_Level04"
    DoCmd.Requery "Description_FeaturesType_Level05"
    DoCmd.Requery "Description_FeaturesCategory_Level01"
End Sub
Private Sub Description_FeaturesCategory_Level01_Click()
    mod_DescriptionFeatures.SetControlVisibility "Description_FeaturesType_Level02", True
    mod_DescriptionFeatures.SetControlVisibility "cmd_Description_FeaturesType_Level02", True
    Description_FeaturesDefinition = "Definition: " & [Description_FeaturesCategory_Level01].[Column](1)
End Sub

'
'   FEATURES TYPE
'   FEATURES TYPE LEVEL ONE
'
Private Sub cmd_Description_FeaturesType_Level01_Click()
    mod_DescriptionFeatures.UpdateFeaturesDescription Me.Description_FeaturesType, Me.Description_FeaturesType_Level01.Value, "Description_FeaturesType"
End Sub
Private Sub Description_FeaturesType_Level01_AfterUpdate()
    DoCmd.Requery "Description_FeaturesType_Level01"
    DoCmd.Requery "Description_FeaturesType_Level02"
    DoCmd.Requery "Description_FeaturesType_Level03"
    DoCmd.Requery "Description_FeaturesType_Level04"
    DoCmd.Requery "Description_FeaturesType_Level05"
    DoCmd.Requery "Description_FeaturesCategory_Level01"
End Sub
Private Sub Description_FeaturesType_Level01_Click()
    Description_FeaturesDefinition = "Definition: " & [Description_FeaturesType_Level01].[Column](1)
End Sub
'
'   FEATURES TYPE LEVEL TWO
'
Private Sub cmd_Description_FeaturesType_Level02_Click()
    mod_DescriptionFeatures.UpdateFeaturesDescription Me.Description_FeaturesType, Me.Description_FeaturesType_Level02.Value, "Description_FeaturesType"
End Sub
Private Sub Description_FeaturesType_Level02_AfterUpdate()
    DoCmd.Requery "Description_FeaturesType_Level01"
    DoCmd.Requery "Description_FeaturesType_Level02"
    DoCmd.Requery "Description_FeaturesType_Level03"
    DoCmd.Requery "Description_FeaturesType_Level04"
    DoCmd.Requery "Description_FeaturesType_Level05"
    DoCmd.Requery "Description_FeaturesCategory_Level01"
End Sub
Private Sub Description_FeaturesType_Level02_Click()
    Description_FeaturesDefinition = "Definition: " & [Description_FeaturesType_Level02].[Column](1)
End Sub
'
'   FEATURES TYPE LEVEL THREE
'
Private Sub cmd_Description_FeaturesType_Level03_Click()
    mod_DescriptionFeatures.UpdateFeaturesDescription Me.Description_FeaturesType, Me.Description_FeaturesType_Level03.Value, "Description_FeaturesType"
End Sub
Private Sub Description_FeaturesType_Level03_AfterUpdate()
    DoCmd.Requery "Description_FeaturesType_Level01"
    DoCmd.Requery "Description_FeaturesType_Level02"
    DoCmd.Requery "Description_FeaturesType_Level03"
    DoCmd.Requery "Description_FeaturesType_Level04"
    DoCmd.Requery "Description_FeaturesType_Level05"
    DoCmd.Requery "Description_FeaturesCategory_Level01"
End Sub
Private Sub Description_FeaturesType_Level03_Click()
    Description_FeaturesDefinition = "Definition: " & [Description_FeaturesType_Level03].[Column](1)
End Sub
'
'   FEATURES TYPE LEVEL FOUR
'
Private Sub cmd_Description_FeaturesType_Level04_Click()
    mod_DescriptionFeatures.UpdateFeaturesDescription Me.Description_FeaturesType, Me.Description_FeaturesType_Level04.Value, "Description_FeaturesType"
End Sub
Private Sub Description_FeaturesType_Level04_AfterUpdate()
    DoCmd.Requery "Description_FeaturesType_Level01"
    DoCmd.Requery "Description_FeaturesType_Level02"
    DoCmd.Requery "Description_FeaturesType_Level03"
    DoCmd.Requery "Description_FeaturesType_Level04"
    DoCmd.Requery "Description_FeaturesType_Level05"
    DoCmd.Requery "Description_FeaturesCategory_Level01"
End Sub
Private Sub Description_FeaturesType_Level04_Click()
    Description_FeaturesDefinition = "Definition: " & [Description_FeaturesType_Level04].[Column](1)
End Sub
'
'   FEATURES TYPE LEVEL FIVE
'
Private Sub cmd_Description_FeaturesType_Level05_Click()
    mod_DescriptionFeatures.UpdateFeaturesDescription Me.Description_FeaturesType, Me.Description_FeaturesType_Level05.Value, "Description_FeaturesType"
End Sub
Private Sub Description_FeaturesType_Level05_AfterUpdate()
    DoCmd.Requery "Description_FeaturesType_Level01"
    DoCmd.Requery "Description_FeaturesType_Level02"
    DoCmd.Requery "Description_FeaturesType_Level03"
    DoCmd.Requery "Description_FeaturesType_Level04"
    DoCmd.Requery "Description_FeaturesType_Level05"
    DoCmd.Requery "Description_FeaturesCategory_Level01"
End Sub
Private Sub Description_FeaturesType_Level05_Click()
    Description_FeaturesDefinition = "Definition: " & [Description_FeaturesType_Level05].[Column](1)
End Sub
'
'   FEATURES TYPE LEVEL ALL LEVELS
'
Private Sub cmd_Description_FeaturesType_AllLevels_Click()
    mod_DescriptionFeatures.UpdateFeaturesDescription Me.Description_FeaturesType, Me.Description_FeaturesType_AllLevels.Value, "Description_FeaturesType"
End Sub
Private Sub Description_FeaturesType_AllLevels_Click()
    Description_FeaturesDefinition = "Definition: " & [Description_FeaturesType_AllLevels].[Column](1)
End Sub
'
'   FEATURES RESET
'
Private Sub cmd_ResetFeaturesDescription_Click()
    ' Resets fields
    Description_FeaturesCategory = " "
    Description_FeaturesCategory_Level01 = " "
    Description_FeaturesType = " "
    Description_FeaturesType_AllLevels = " "
    Description_FeaturesType_Level01 = " "
    Description_FeaturesType_Level02 = " "
    Description_FeaturesType_Level03 = " "
    Description_FeaturesType_Level04 = " "
    Description_FeaturesType_Level05 = " "
    Description_FeaturesDefinition = " "
End Sub
'
'  ------ FOOTERS SECTION ------------------------------------------------------------------------------------
'
'  FILTERS
'
Private Sub cmdRemoveFilter_Click()
    On Error Resume Next
    DoCmd.RunCommand acCmdRemoveAllFilters
    If Err.Number = 0 Then
        Me.FilterOn = False
        MsgBox "All filters have been removed."
    Else
        MsgBox "Error " & Err.Number & ": " & Err.Description, vbExclamation, "Filter Removal Error"
    End If
End Sub
Private Sub cmd_GoToRecord_Click()
    Dim recordID As Variant
    Dim rs As Recordset

    recordID = Me.txt_GoTo.Value

    If IsNumeric(recordID) Then
        Set rs = Me.RecordsetClone
        rs.FindFirst "[ID_FindID] = " & recordID

        If rs.NoMatch Then
            MsgBox "Record not found.", vbInformation, "Search Result"
        Else
            Me.Bookmark = rs.Bookmark
        End If
    Else
        MsgBox "Please enter a valid numeric ID.", vbInformation, "Invalid Input"
    End If

    If Not rs Is Nothing Then
        Set rs = Nothing
    End If
End Sub
'
'  RECORD SELECTORS SECTION
'
Private Sub Form_Current()
    'Displays Record Number
    On Error GoTo ErrorHandler

    ' Update record count display
    Me.txtRecs = "Record " & Me.CurrentRecord & " of " & Me.RecordsetClone.RecordCount

    Exit Sub

ErrorHandler:
    ' Handle potential errors, for example when recordset is not available
    Me.txtRecs = "Record information unavailable"
    ' Optionally, log error or display a message box
End Sub
Private Sub cmdFirst_Click()
    'Go to first record
    If Not (Me.Recordset.BOF) Then
        Me.Recordset.MoveFirst
    End If
End Sub
Private Sub cmdLast_Click()
    'Go to last record
    If Not (Me.Recordset.EOF) Then
        Me.Recordset.MoveLast
    End If
End Sub
Private Sub cmdNext_Click()
    'Go to next record
    If Not (Me.Recordset.EOF) Then
        Me.Recordset.MoveNext
        If Me.Recordset.EOF Then Me.Recordset.MoveLast
    End If
End Sub
Private Sub cmdPrevious_Click()
    'Go to previous record
    If Not (Me.Recordset.BOF) Then
        Me.Recordset.MovePrevious
        If Me.Recordset.BOF Then Me.Recordset.MoveFirst
    End If
End Sub
Private Sub cmdNew_Click()
    'Go to new record
    On Error GoTo ErrorHandler
    DoCmd.RunCommand acCmdRecordsGoToNew
    Exit Sub

ErrorHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Error"
End Sub
