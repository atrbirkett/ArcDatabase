VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frm_BuildingRecords"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
'  BEGINING OF MASONRY SECTION
'
Private Sub cmd_Description_MASONRY_02_Click()
' Adds selection to MASONRY/Category from MASONRY/Category
     Description_Masonry = Description_Masonry & "; " & Description_Masonry_02
End Sub
Private Sub cmd_Description_MASONRY_FL_Click()
' Adds selection to MASONRY/Category from MASONRY/Category
     Description_Masonry = Description_Masonry & "; " & Description_Masonry_FL
End Sub
Private Sub Description_MasonryCategory_01_AfterUpdate()
    DoCmd.Requery "Description_Masonry_01"
    DoCmd.Requery "Description_Masonry_02"
End Sub
Private Sub Description_Masonry_01_AfterUpdate()
    DoCmd.Requery "Description_Masonry_02"
End Sub
Private Sub cmd_ResetMasonryDescription_Click()
' Resets and clears all fields and hides all layers of item
' description search when a new category is chosen
    Description_MasonryCategory_01 = ""
    Description_Masonry = ""
    Description_Masonry_01 = ""
    Description_MasoneyDeff = ""
    Description_Masonry_02 = ""
End Sub
Private Sub Description_MASONRYCategory_01_Click()
' Adds selection to Type/Category from Item/Category and unhides type
    Description_Masonry_01 = ""
    Description_Masonry_02 = ""
    Description_Masonry_FL = ""
    Description_MasonryDeff = "Definition: " & [Description_MasonryCategory_01].[Column](1)
End Sub
Private Sub Description_MASONRY_02_Click()
    Description_MasonryDeff = "Definition: " & [Description_Masonry_02].[Column](1)
End Sub
Private Sub Description_MASONRY_FL_Click()
    Description_MasonryDeff = "Definition: " & [Description_Masonry_FL].[Column](1)
End Sub
Private Sub Description_MASONRY_01_Click()
    Description_MasonryDeff = "Definition: " & [Description_Masonry_01].[Column](1)
End Sub

Private Sub cmd_Description_Masonry_01_Click()
     Description_Masonry = Description_Masonry & "; " & Description_Masonry_01
End Sub
Private Sub cmd_Description_MasonryCategory_01_Click()
     Description_Masonry = Description_Masonry & "; " & Description_MasonryCategory_01
End Sub
'
'  END OF MASONRYY SECTION
'
Private Sub cmdBackup_Click()
    On Error GoTo cmdBackup_Click_Err
        DoCmd.RunCommand acCmdExportXML
cmdBackup_Click_Exit:
        Exit Sub
cmdBackup_Click_Err:
    MsgBox Error$
    Resume cmdBackup_Click_Exit
End Sub

Private Sub cmdRemoveFilter_Click()
      On Error GoTo NoFilter
    DoCmd.RunCommand acCmdRemoveAllFilters
    Me.Form.FilterOn = False
NoFilter:
    MsgBox "You now have no filters set."
End Sub

Private Sub Form_Current()
   
    Me.txtRecs = "Record" & " " & Form.CurrentRecord & " " & "of" & " " & Form.RecordsetClone.RecordCount

End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)
        
    DoCmd.SetWarnings False
    DoCmd.RunSQL "INSERT INTO bak_MasonryRecords SELECT ID_MasonryNumber,ID_Trench,ID_Room_List,Location_X,Location_Y,Location_Address,Location_Grid,Mesure_Facing,Mesure_MaxH,Mesure_LowestH,Mesure_Length,Measure_WidthMax,Measure_WidthMed,Measure_WidthMin,Mesure_BedMax,Mesure_HeaderMax,Mesure_StretcherMax,Mesure_BedMed,Mesure_HeaderMed,Mesure_StretcherMed,Mesure_BedMin,Mesure_HeaderMin,Mesure_StretcherMin,Material_Type,Material_Finish,Coursing_Height,Coursing_Type,Coursing_Description,Bonding_Description,Record_Notes,Harris_EqualWith_List,Harris_EarlierThan_List,Harris_ContemporaryWith_List,Harris_LaterThan_List,Samples_WorkedStonesSpecial_List,Samples_WorkedStonesRegular_List FROM tbl_MasonryRecords " & _
        "WHERE UID=" & UID
    DoCmd.SetWarnings True
    
End Sub

Private Sub ID_Room_AfterUpdate()
ID_Room_List = ID_Room
End Sub
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

Private Sub cmdSetAssRecords_Click()

    Samples_WorkedStonesSpecial_List = Samples_WorkedStonesSpecial
    Samples_WorkedStonesRegular_List = Samples_WorkedStonesRegular

End Sub
Private Sub txt_GoTo_AfterUpdate()
    Me.Form.Filter = "ID_MasonryNumber=" & Me.txt_GoTo
    Me.Form.FilterOn = True
End Sub
Private Sub cmdSetHarrisRecords_Click()

    Harris_EarlierThan_List = Harris_EarlierThan
    Harris_ContemporaryWith_List = Harris_ContemporaryWith
    Harris_LaterThan_List = Harris_LaterThan
    Harris_EqualWith_List = Harris_EqualWith

End Sub

Private Sub Ctl00_AfterUpdate()
    ID_Room_List = ID_Room
End Sub