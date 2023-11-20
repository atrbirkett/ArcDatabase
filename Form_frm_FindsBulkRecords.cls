VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frm_FindsBulkRecords"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub AccessionID_AfterUpdate()
 Me.ID_AccessionNumber = "2022.HARP/" & ID_BoxID.Value & "BF" & Right("0000" & CStr(ID_BagID), 4)
End Sub

Private Sub AccessionID_Click()
    Me.ID_AccessionNumber = "2022.HARP/" & ID_BoxID.Value & "BF" & Right("0000" & CStr(ID_BagID), 4)
End Sub
Private Sub BoxID_AfterUpdate()
 Me.AccessionID = "2022.HARP/" & ID_BoxID.Value & "/BF" & Right("0000" & CStr(ID_BagID), 4)
End Sub
Private Sub cmdAssemblage_Click()
 DoCmd.OpenForm "frm_Assemblages", acNormal, "", "", acFormEdit, acNormal
End Sub
Private Sub cmdFilter_Click()
    Me.Form.Filter = "ID_BagID=" & Me.txt_GoTo
    Me.Form.FilterOn = True
End Sub

Private Sub cmdRemoveFilter_Click()
      On Error GoTo NoFilter
    DoCmd.RunCommand acCmdRemoveAllFilters
    Me.Form.FilterOn = False
NoFilter:
    MsgBox "You now have no filters set."
End Sub
Private Sub Dating_Period_AfterUpdate()
Dating_MinYear = [Dating_Period].[Column](2)
Dating_MaxYear = [Dating_Period].[Column](3)
End Sub

Private Sub Form_AfterUpdate()
    Me.Description_FeaturesType_Level01.Visible = False
    Me.cmd_Description_MaterialClass_Level01.Visible = False
    Me.Description_FeaturesType_Level02.Visible = False
    Me.cmd_Description_MaterialClass_Level02.Visible = False
    Me.Description_FeaturesType_Level03.Visible = False
    Me.cmd_Description_MaterialClass_Level03.Visible = False
    Me.Description_FeaturesType_Level04.Visible = False
    Me.cmd_Description_FeaturesType_Level04.Visible = False
    Me.Description_FeaturesType_Level05.Visible = False
    Me.cmd_Description_FeaturesType_Level05.Visible = False
    Me.Description_MaterialClass_Level01.Visible = False
    Me.cmd_Description_FeaturesType_Level01.Visible = False
    Me.Description_MaterialClass_Level02.Visible = False
    Me.cmd_Description_FeaturesType_Level02.Visible = False
    Me.Description_MaterialClass_Level03.Visible = False
    Me.cmd_Description_FeaturesType_Level03.Visible = False

End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)
    DoCmd.SetWarnings False
    DoCmd.RunSQL "INSERT INTO bak_FindsBulkRecords SELECT ID_BagID,ID_AssemblageID,ID_Unstratified,ID_Trench,ID_Context,Location_Grid,Count_NumberofType,Count_CountIsEstimation,Count_Weight,Material_Classification,Material_Type,Find_Classificiation,Find_Type,Find_SpecificName,Dating_Period,Dating_MinYear,Dating_MaxYear,Description_Form,Description_Colour,Description_Decoration,Description_Interp,Description_Notes FROM tbl_FindsBulkRecords " &
        "WHERE UID=" & UID
    DoCmd.SetWarnings True

End Sub

Private Sub Form_Load()
    Me.Description_FeaturesType_Level01.Visible = False
    Me.cmd_Description_MaterialClass_Level01.Visible = False
    Me.Description_FeaturesType_Level02.Visible = False
    Me.cmd_Description_MaterialClass_Level02.Visible = False
    Me.Description_FeaturesType_Level03.Visible = False
    Me.cmd_Description_MaterialClass_Level03.Visible = False
    Me.Description_FeaturesType_Level04.Visible = False
    Me.cmd_Description_FeaturesType_Level04.Visible = False
    Me.Description_FeaturesType_Level05.Visible = False
    Me.cmd_Description_FeaturesType_Level05.Visible = False
    Me.Description_MaterialClass_Level01.Visible = False
    Me.cmd_Description_FeaturesType_Level01.Visible = False
    Me.Description_MaterialClass_Level02.Visible = False
    Me.cmd_Description_FeaturesType_Level02.Visible = False
    Me.Description_MaterialClass_Level03.Visible = False
    Me.cmd_Description_FeaturesType_Level03.Visible = False
End Sub

Private Sub ID_BagIDField_AfterUpdate()
    On Error GoTo ErrorHandler
    Me.ID_AccessionNumber = "2022.HARP/" & ID_BoxID.Value & "BF" & Right("0000" & CStr(ID_BagID), 4)
    Exit Sub
ErrorHandler:
    MsgBox "An error occurred: " & Err.Description
End Sub
Private Sub Form_Current()
    Me.txtRecs = "Record" & " " & Form.CurrentRecord & " of " & Form.RecordsetClone.RecordCount
End Sub
'
'   BEGINING OF MATERIAL SECTION
'
'   MATERIAL TYPE
'   MATERIAL TYPE LEVEL ONE
'
Private Sub cmd_Description_MaterialType_Level01_Click()
    mod_DescriptionMaterial.UpdateMaterialDescription Description_MaterialType, Description_MaterialType_Level01, "Description_MaterialType"
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
    mod_DescriptionMaterial.UpdateMaterialDescription Description_MaterialClass, Description_MaterialClass_Level01, "Description_MaterialClass"
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
    mod_DescriptionMaterial.UpdateMaterialDescription Description_MaterialClass, Description_MaterialClass_Level02, "Description_MaterialClass"
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
    mod_DescriptionMaterial.UpdateMaterialDescription Description_MaterialClass, Description_MaterialClass_Level03, "Description_MaterialClass"
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
    mod_DescriptionMaterial.UpdateMaterialDescription Description_MaterialClass, Description_MaterialClass_AllLevels, "Description_MaterialClass"
End Sub
Private Sub Description_MaterialClass_AllLevels_Click()
    Description_MatDefinition = "Definition: " & [Description_MaterialClass_AllLevels].[Column](1)
End Sub
'
'   MATERIAL RESET
'
Private Sub cmd_ResetMatDescription_Click()
    ' Resets fields
    Description_MaterialType = ""
    Description_MaterialType_Level01 = ""
    Description_MaterialClass = ""
    Description_MaterialClass_AllLevels = ""
    Description_MaterialClass_Level01 = ""
    Description_MaterialClass_Level02 = ""
    Description_MaterialClass_Level03 = ""
    Description_MatDefinition = ""

    ' Update visibility of controls
    mod_DescriptionMaterial.SetControlVisibility "lab_SearchAllLevels", False
    mod_DescriptionMaterial.SetControlVisibility "Description_MaterialClass_AllLevels", False
    mod_DescriptionMaterial.SetControlVisibility "cmd_Description_MaterialClass_AllLevels", False
    mod_DescriptionMaterial.SetControlVisibility "Description_MaterialClass_Level01", False
    mod_DescriptionMaterial.SetControlVisibility "cmd_Description_MaterialClass_Level01", False
    mod_DescriptionMaterial.SetControlVisibility "Description_MaterialClass_Level02", False
    mod_DescriptionMaterial.SetControlVisibility "cmd_Description_MaterialClass_Level02", False
    mod_DescriptionMaterial.SetControlVisibility "Description_MaterialClass_Level03", False
    mod_DescriptionMaterial.SetControlVisibility "cmd_Description_MaterialClass_Level03", False
End Sub
'
'  END OF MATERIAL SECTION
'
'
'  BEGINING OF FEATURES SECTION
'
Private Sub cmd_Description_FeaturesCategory_01_Click()
    If IsNull(Description_FeaturesCategory) Then Description_FeaturesCategory = Description_FeaturesCategory_01 Else : Description_FeaturesCategory = Description_FeaturesCategory & "; " & Description_FeaturesCategory_01
End Sub
Private Sub Description_FeaturesCategory_01_DblClick(Cancel As Integer)
    If IsNull(Description_FeaturesCategory) Then Description_FeaturesCategory = Description_FeaturesCategory_01 Else : Description_FeaturesCategory = Description_FeaturesCategory & "; " & Description_FeaturesCategory_01
End Sub
Private Sub Description_FeaturesType_Level01_DblClick(Cancel As Integer)
    ' Adds selection to Features/Category from Features/Category
    If IsNull(Description_FeaturesType) Then Description_FeaturesType = Description_FeaturesType_Level01 Else : Description_FeaturesType = Description_FeaturesType & "; " & Description_FeaturesType_Level01
End Sub
Private Sub cmd_Description_FeaturesType_Level01_Click()
    ' Adds selection to Features/Category from Features/Category
    If IsNull(Description_FeaturesType) Then Description_FeaturesType = Description_FeaturesType_Level01 Else : Description_FeaturesType = Description_FeaturesType & "; " & Description_FeaturesType_Level01
End Sub
Private Sub Description_FeaturesType_Level02_DblClick(Cancel As Integer)
    ' Adds selection to Features/Category from Features/Category
    Description_FeaturesType = Description_FeaturesType & "; " & Description_FeaturesType_Level02
End Sub
Private Sub cmd_Description_FeaturesType_Level02_Click()
    ' Adds selection to Features/Category from Features/Category
    Description_FeaturesType = Description_FeaturesType & "; " & Description_FeaturesType_Level02
End Sub
Private Sub Description_FeaturesType_Level03_DblClick(Cancel As Integer)
    ' Adds selection to Features/Category from Features/Category
    Description_FeaturesType = Description_FeaturesType & "; " & Description_FeaturesType_Level03
End Sub
Private Sub cmd_Description_FeaturesType_Level03_Click()
    ' Adds selection to Features/Category from Features/Category
    Description_FeaturesType = Description_FeaturesType & "; " & Description_FeaturesType_Level03
End Sub
Private Sub cmd_Description_FeaturesType_Level04_Click()
    ' Adds selection to Features/Category from Features/Category
    Description_FeaturesType = Description_FeaturesType & "; " & Description_FeaturesType_Level04
End Sub
Private Sub Description_FeaturesType_DblClick(Cancel As Integer)
    ' Adds selection to Features/Category from Features/Category
    Description_FeaturesType = Description_FeaturesType & "; " & Description_FeaturesType_Level04
End Sub
Private Sub cmd_Description_FeaturesType_Level05_Click()
    ' Adds selection to Features/Category from Features/Category
    Description_FeaturesType = Description_FeaturesType & "; " & Description_FeaturesType_Level05
End Sub
Private Sub Description_FeaturesType_Level05_DblClick(Cancel As Integer)
    ' Adds selection to Features/Category from Features/Category
    Description_FeaturesType = Description_FeaturesType & "; " & Description_FeaturesType_Level05
End Sub
Private Sub cmd_Description_FeaturesType_AllLevels_Click()
    ' Adds selection to Type/Category from Item/Category
    If IsNull(Description_FeaturesCategory) Then Description_FeaturesType = Description_FeaturesType_AllLevels Else : Description_FeaturesType = Description_FeaturesType & "; " & Description_FeaturesType_AllLevels
End Sub
Private Sub Description_FeaturesType_AllLevels_DblClick(Cancel As Integer)
    ' Adds selection to Type/Category from Item/Category
    If IsNull(Description_FeaturesCategory) Then Description_FeaturesType = Description_FeaturesType_AllLevels Else : Description_FeaturesType = Description_FeaturesType & "; " & Description_FeaturesType_AllLevels
End Sub

Private Sub cmd_Reset_FeaturesDescription_Click()
    ' Resets and clears all fields and hides all layers of item
    ' description search when a new category is chosen
    Description_FeaturesCategory_01 = ""
    Description_FeaturesCategory = ""
    Description_FeaturesType = ""
    Description_FeaturesType_AllLevels = ""
    Description_FeaturesType_Level01 = ""
    Description_FeaturesTypeDeff = ""
    DoCmd.SetProperty "Description_FeaturesType_Level01", acPropertyVisible, "0"
    DoCmd.SetProperty "cmd_Description_FeaturesType_Level01", acPropertyVisible, "0"
    Description_FeaturesType_Level02 = ""
    DoCmd.SetProperty "Description_FeaturesType_Level02", acPropertyVisible, "0"
    DoCmd.SetProperty "cmd_Description_FeaturesType_Level02", acPropertyVisible, "0"
    Description_FeaturesType_Level03 = ""
    DoCmd.SetProperty "Description_FeaturesType_Level03", acPropertyVisible, "0"
    DoCmd.SetProperty "cmd_Description_FeaturesType_Level03", acPropertyVisible, "0"
        Description_FeaturesType_Level04 = ""
    DoCmd.SetProperty "Description_FeaturesType_Level04", acPropertyVisible, "0"
    DoCmd.SetProperty "cmd_Description_FeaturesType_Level04", acPropertyVisible, "0"
        Description_FeaturesType_Level05 = ""
    DoCmd.SetProperty "Description_FeaturesType_Level05", acPropertyVisible, "0"
    DoCmd.SetProperty "cmd_Description_FeaturesType_Level05", acPropertyVisible, "0"
    Description_FeaturesTypeDeff = ""
End Sub
Private Sub Description_FeaturesType_Level01_AfterUpdate()
    DoCmd.Requery "Description_FeaturesType_Level01"
    DoCmd.Requery "Description_FeaturesType_Level02"
    DoCmd.Requery "Description_FeaturesType_Level03"
    DoCmd.Requery "Description_FeaturesType_Level05"
    DoCmd.Requery "Description_FeaturesType_Level04"
    DoCmd.Requery "Description_FeaturesCategory_01"
End Sub

Private Sub Description_FeaturesType_Level02_AfterUpdate()
    DoCmd.Requery "Description_FeaturesType_Level01"
    DoCmd.Requery "Description_FeaturesType_Level02"
    DoCmd.Requery "Description_FeaturesType_Level03"
        DoCmd.Requery "Description_FeaturesType_Level05"
    DoCmd.Requery "Description_FeaturesType_Level04"
    DoCmd.Requery "Description_FeaturesCategory_01"
End Sub
Private Sub Description_FeaturesType_Level03_AfterUpdate()
    DoCmd.Requery "Description_FeaturesType_Level01"
    DoCmd.Requery "Description_FeaturesType_Level02"
    DoCmd.Requery "Description_FeaturesType_Level03"
        DoCmd.Requery "Description_FeaturesType_Level05"
    DoCmd.Requery "Description_FeaturesType_Level04"
    DoCmd.Requery "Description_FeaturesCategory_01"
End Sub
Private Sub Description_FeaturesCategory_01_AfterUpdate()
    DoCmd.Requery "Description_FeaturesType_Level01"
    DoCmd.Requery "Description_FeaturesType_Level02"
    DoCmd.Requery "Description_FeaturesType_Level03"
        DoCmd.Requery "Description_FeaturesType_Level05"
    DoCmd.Requery "Description_FeaturesType_Level04"
    DoCmd.Requery "Description_FeaturesCategory_01"
End Sub
Private Sub Description_FeaturesCategory_01_Click()
    ' Adds selection to Type/Category from Item/Category and unhides type
    Description_FeaturesType_Level02 = ""
    DoCmd.SetProperty "Description_FeaturesType_Level02", acPropertyVisible, "0"
    DoCmd.SetProperty "cmd_Description_FeaturesType_Level02", acPropertyVisible, "0"
    Description_FeaturesType_Level03 = ""
    DoCmd.SetProperty "Description_FeaturesType_Level03", acPropertyVisible, "0"
    DoCmd.SetProperty "cmd_Description_FeaturesType_Level03", acPropertyVisible, "0"
        Description_FeaturesType_Level04 = ""
    DoCmd.SetProperty "Description_FeaturesType_Level04", acPropertyVisible, "0"
    DoCmd.SetProperty "cmd_Description_FeaturesType_Level04", acPropertyVisible, "0"
            Description_FeaturesType_Level05 = ""
    DoCmd.SetProperty "Description_FeaturesType_Level05", acPropertyVisible, "0"
    DoCmd.SetProperty "cmd_Description_FeaturesType_Level05", acPropertyVisible, "0"
    ' Adds selection to Type/Category from Item/Category and unhides type
    Description_FeaturesType_AllLevels = ""
    DoCmd.SetProperty "Description_FeaturesType_AllLevels", acPropertyVisible, "-1"
    DoCmd.SetProperty "cmd_Description_FeaturesType_AllLevels", acPropertyVisible, "-1"
    Description_FeaturesType_Level01 = ""
    DoCmd.SetProperty "Description_FeaturesType_Level01", acPropertyVisible, "-1"
    DoCmd.SetProperty "cmd_Description_FeaturesType_Level01", acPropertyVisible, "-1"
    Description_FeaturesTypeDeff = "Definition: " & [Description_FeaturesCategory_01].[Column](1)
End Sub
Private Sub Description_FeaturesType_Level01_Click()
    ' Unhides sub-layers of search Featuress
    DoCmd.SetProperty "Description_FeaturesType_Level02", acPropertyVisible, "-1"
    DoCmd.SetProperty "cmd_Description_FeaturesType_Level02", acPropertyVisible, "-1"
    Description_FeaturesTypeDeff = "Definition: " & [Description_FeaturesType_Level01].[Column](1)
End Sub
Private Sub Description_FeaturesType_Level02_Click()
    ' Unhides sub-layers of search Featuress
    DoCmd.SetProperty "Description_FeaturesType_Level03", acPropertyVisible, "-1"
    DoCmd.SetProperty "cmd_Description_FeaturesType_Level03", acPropertyVisible, "-1"
    Description_FeaturesTypeDeff = "Definition: " & [Description_FeaturesType_Level02].[Column](1)
End Sub
Private Sub Description_FeaturesType_Level03_Click()
    ' Unhides sub-layers of search Featuress
    DoCmd.SetProperty "Description_FeaturesType_Level04", acPropertyVisible, "-1"
    DoCmd.SetProperty "cmd_Description_FeaturesType_Level04", acPropertyVisible, "-1"
    Description_FeaturesTypeDeff = "Definition: " & [Description_FeaturesType_Level03].[Column](1)
End Sub
Private Sub Description_FeaturesType_Level04_Click()
    ' Unhides sub-layers of search Featuress
    DoCmd.SetProperty "Description_FeaturesType_Level05", acPropertyVisible, "-1"
    DoCmd.SetProperty "cmd_Description_FeaturesType_Level05", acPropertyVisible, "-1"
    Description_FeaturesTypeDeff = "Definition: " & [Description_FeaturesType_Level04].[Column](1)
End Sub
Private Sub Description_FeaturesType_Level05_Click()
    ' Unhides sub-layers of search Featuress
    Description_FeaturesTypeDeff = "Definition: " & [Description_FeaturesType_Level05].[Column](1)
End Sub
Private Sub Description_FeaturesType_AllLevels_Click()
    Description_FeaturesTypeDeff = "Definition: " & [Description_FeaturesType_AllLevels].[Column](1) & " (" & [Description_FeaturesType_AllLevels].[Column](2) & ")"
End Sub

'
'  END OF FEATURES SECTION
'
Private Sub cmdFirst_Click()
    On Error Resume Next
    Me.Recordset.MoveFirst
End Sub

Private Sub cmdLast_Click()
    On Error Resume Next
    Me.Recordset.MoveLast
End Sub

Private Sub cmdNext_Click()
    On Error Resume Next
    Me.Recordset.MoveNext
End Sub

Private Sub cmdPrevious_Click()
    On Error Resume Next
    Me.Recordset.MovePrevious
End Sub

Private Sub cmdNew_Click()
    On Error Resume Next
    DoCmd.RunCommand acCmdRecordsGoToNew
End Sub
