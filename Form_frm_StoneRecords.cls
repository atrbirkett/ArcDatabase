VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frm_StoneRecords"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'
'  BEGINING OF PURPOSE SECTION
'
Private Sub cmd_Description_PURPOSECategory_01_Click()
      Description_Purpose = Description_Purpose & "; " & Description_PurposeCategory_01
End Sub
Private Sub cmd_Description_PURPOSE_01_Click()
' Adds selection to PURPOSE/Category from PURPOSE/Category
     Description_Purpose = Description_Purpose & "; " & Description_Purpose_01
End Sub

Private Sub Form_AfterUpdate()
    DoCmd.SetWarnings False
    DoCmd.RunSQL "INSERT INTO bak_WorkedStoneRecords SELECT ID_StoneNumber,ID_Building,ID_Masonry,ID_Trench,ID_Context,Check_Retained,Location_Address,Location_Grid,Location_X,Location_Y,Mesure_Weight,Mesure_Depth,Mesure_Width,Mesure_Length,Mesure_Diam,Description_Dimensions,Material_Type,Description_Material,Description_Finish,Description_Purpose,Description_Characteristics,Description_Condition,Description_Comments FROM tbl_WorkedStoneRecords " & _
        "WHERE UID=" & UID
    DoCmd.SetWarnings True
End Sub

Private Sub txt_GoTo_AfterUpdate()
    Me.Form.Filter = "ID_StoneNumber=" & Me.txt_GoTo
    Me.Form.FilterOn = True
End Sub
Private Sub cmd_Description_PURPOSE_02_Click()
' Adds selection to PURPOSE/Category from PURPOSE/Category
     Description_Purpose = Description_Purpose & "; " & Description_Purpose_02
End Sub

Private Sub cmd_Description_PURPOSE_03_Click()
' Adds selection to PURPOSE/Category from PURPOSE/Category
     Description_Purpose = Description_Purpose & "; " & Description_Purpose_03
End Sub
Private Sub cmd_Description_PURPOSE_04_Click()
' Adds selection to PURPOSE/Category from PURPOSE/Category
     Description_Purpose = Description_Purpose & "; " & Description_Purpose_04
End Sub

Private Sub cmd_Description_PURPOSE_05_Click()
' Adds selection to PURPOSE/Category from PURPOSE/Category
     Description_Purpose = Description_Purpose & "; " & Description_Purpose_05
End Sub
Private Sub cmd_Description_PURPOSE_FL_Click()
' Adds selection to Type/Category from Item/Category
    Description_Purpose = Description_Purpose & "; " & Description_Purpose_FL
End Sub


Private Sub cmd_Reset_PURPOSEDescription_Click()
' Resets and clears all fields and hides all layers of item
' description search when a new category is chosen
    Description_PurposeCategory_01 = ""
    Description_Purpose = ""
    Description_Purpose_FL = ""
    Description_Purpose_01 = ""
    Description_PurposeDeff = ""
    DoCmd.SetProperty "Description_PURPOSE_01", acPropertyVisible, "0"
    DoCmd.SetProperty "cmd_Description_PURPOSE_01", acPropertyVisible, "0"
    Description_Purpose_02 = ""
    DoCmd.SetProperty "Description_PURPOSE_02", acPropertyVisible, "0"
    DoCmd.SetProperty "cmd_Description_PURPOSE_02", acPropertyVisible, "0"
    Description_Purpose_03 = ""
    DoCmd.SetProperty "Description_PURPOSE_03", acPropertyVisible, "0"
    DoCmd.SetProperty "cmd_Description_PURPOSE_03", acPropertyVisible, "0"
        Description_Purpose_04 = ""
    DoCmd.SetProperty "Description_PURPOSE_03", acPropertyVisible, "0"
    DoCmd.SetProperty "cmd_Description_PURPOSE_03", acPropertyVisible, "0"
        Description_Purpose_05 = ""
    DoCmd.SetProperty "Description_PURPOSE_03", acPropertyVisible, "0"
    DoCmd.SetProperty "cmd_Description_PURPOSE_03", acPropertyVisible, "0"
    Description_PurposeDeff = ""
End Sub

Private Sub Description_PURPOSE_01_AfterUpdate()
    DoCmd.Requery "Description_PURPOSE_01"
    DoCmd.Requery "Description_PURPOSE_02"
    DoCmd.Requery "Description_PURPOSE_03"
    DoCmd.Requery "Description_PURPOSE_04"
    DoCmd.Requery "Description_PURPOSE_05"
    DoCmd.Requery "Description_PURPOSECategory_01"
End Sub

Private Sub Description_PURPOSE_02_AfterUpdate()
    DoCmd.Requery "Description_PURPOSE_01"
    DoCmd.Requery "Description_PURPOSE_02"
    DoCmd.Requery "Description_PURPOSE_03"
    DoCmd.Requery "Description_PURPOSE_04"
    DoCmd.Requery "Description_PURPOSE_05"
    DoCmd.Requery "Description_PURPOSECategory_01"
End Sub
Private Sub Description_PURPOSE_03_AfterUpdate()
    DoCmd.Requery "Description_PURPOSE_01"
    DoCmd.Requery "Description_PURPOSE_02"
    DoCmd.Requery "Description_PURPOSE_03"
    DoCmd.Requery "Description_PURPOSE_04"
    DoCmd.Requery "Description_PURPOSE_05"
    DoCmd.Requery "Description_PURPOSECategory_01"
End Sub

Private Sub Description_Purpose_04_AfterUpdate()
    DoCmd.Requery "Description_PURPOSE_01"
    DoCmd.Requery "Description_PURPOSE_02"
    DoCmd.Requery "Description_PURPOSE_03"
    DoCmd.Requery "Description_PURPOSE_04"
    DoCmd.Requery "Description_PURPOSE_05"
    DoCmd.Requery "Description_PURPOSECategory_01"
End Sub

Private Sub Description_Purpose_05_AfterUpdate()
    DoCmd.Requery "Description_PURPOSE_01"
    DoCmd.Requery "Description_PURPOSE_02"
    DoCmd.Requery "Description_PURPOSE_03"
    DoCmd.Requery "Description_PURPOSE_04"
    DoCmd.Requery "Description_PURPOSE_05"
    DoCmd.Requery "Description_PURPOSECategory_01"
End Sub

Private Sub Description_PURPOSECategory_01_AfterUpdate()
    DoCmd.Requery "Description_PURPOSE_01"
    DoCmd.Requery "Description_PURPOSE_02"
    DoCmd.Requery "Description_PURPOSE_03"
    DoCmd.Requery "Description_PURPOSECategory_01"
End Sub
Private Sub Description_PURPOSECategory_01_Click()
' Adds selection to Type/Category from Item/Category and unhides type
    Description_Purpose_02 = ""
    DoCmd.SetProperty "Description_PURPOSE_02", acPropertyVisible, "0"
    DoCmd.SetProperty "cmd_Description_PURPOSE_02", acPropertyVisible, "0"
    Description_Purpose_03 = ""
    DoCmd.SetProperty "Description_PURPOSE_03", acPropertyVisible, "0"
    DoCmd.SetProperty "cmd_Description_PURPOSE_03", acPropertyVisible, "0"
    ' Adds selection to Type/Category from Item/Category and unhides type
    Description_Purpose_FL = ""
    DoCmd.SetProperty "Description_PURPOSE_FL", acPropertyVisible, "-1"
    DoCmd.SetProperty "cmd_Description_PURPOSE_FL", acPropertyVisible, "-1"
    Description_Purpose_01 = ""
    DoCmd.SetProperty "Description_PURPOSE_01", acPropertyVisible, "-1"
    DoCmd.SetProperty "cmd_Description_PURPOSE_01", acPropertyVisible, "-1"
        Description_Purpose_04 = ""
    DoCmd.SetProperty "Description_PURPOSE_04", acPropertyVisible, "0"
    DoCmd.SetProperty "cmd_Description_PURPOSE_04", acPropertyVisible, "0"
        Description_Purpose_05 = ""
    DoCmd.SetProperty "Description_PURPOSE_05", acPropertyVisible, "0"
    DoCmd.SetProperty "cmd_Description_PURPOSE_05", acPropertyVisible, "0"
    Description_PurposeDeff = "Deffinition: " & [Description_PurposeCategory_01].[Column](1)
End Sub
Private Sub Description_PURPOSE_01_Click()
' Unhides sub-layers of search PURPOSEs
    DoCmd.SetProperty "Description_PURPOSE_02", acPropertyVisible, "-1"
    DoCmd.SetProperty "cmd_Description_PURPOSE_02", acPropertyVisible, "-1"
    Description_PurposeDeff = "Deffinition: " & [Description_Purpose_01].[Column](1)
End Sub
Private Sub Description_PURPOSE_02_Click()
' Unhides sub-layers of search PURPOSEs
    DoCmd.SetProperty "Description_PURPOSE_03", acPropertyVisible, "-1"
    DoCmd.SetProperty "cmd_Description_PURPOSE_03", acPropertyVisible, "-1"
    Description_PurposeDeff = "Deffinition: " & [Description_Purpose_02].[Column](1)
End Sub
Private Sub Description_PURPOSE_03_Click()
' Unhides sub-layers of search PURPOSEs
    DoCmd.SetProperty "Description_PURPOSE_04", acPropertyVisible, "-1"
    DoCmd.SetProperty "cmd_Description_PURPOSE_04", acPropertyVisible, "-1"
    Description_PurposeDeff = "Deffinition: " & [Description_Purpose_03].[Column](1)
End Sub
Private Sub Description_PURPOSE_04_Click()
' Unhides sub-layers of search PURPOSEs
    DoCmd.SetProperty "Description_PURPOSE_05", acPropertyVisible, "-1"
    DoCmd.SetProperty "cmd_Description_PURPOSE_05", acPropertyVisible, "-1"
    Description_PurposeDeff = "Deffinition: " & [Description_Purpose_04].[Column](1)
End Sub
Private Sub Description_PURPOSE_05_Click()
' Unhides sub-layers of search PURPOSEs
    Description_PurposeDeff = "Deffinition: " & [Description_Purpose_05].[Column](1)
End Sub
Private Sub Description_PURPOSE_FL_Click()
      Description_PurposeDeff = "Deffinition: " & [Description_Purpose_FL].[Column](1) & " (" & [Description_Purpose_FL].[Column](2) & ")"
End Sub

'
'  END OF PURPOSE SECTION
'
'
'  BEGINING OF MATERIAL SECTION
'
Private Sub cmd_Description_Material_01_Click()
' Adds selection to Material/Category from Material/Category
     Description_Material = Description_Material & "; " & Description_Material_01
     DoCmd.Requery "Description_Material"
End Sub

Private Sub cmd_Description_Material_02_Click()
' Adds selection to Material/Category from Material/Category
     Description_Material = Description_Material & "; " & Description_Material_02
     DoCmd.Requery "Description_Material"
End Sub

Private Sub cmd_Description_Material_03_Click()
' Adds selection to Material/Category from Material/Category
     Description_Material = Description_Material & "; " & Description_Material_03
     DoCmd.Requery "Description_Material"
End Sub


Private Sub cmd_Description_Material_FL_Click()
' Adds selection to Type/Category from Item/Category
    Description_Material = Description_Material & "; " & Description_Material_FL
    DoCmd.Requery "Description_Material"
End Sub

Private Sub cmd_Description_MaterialCategory_01_Click()
    Description_Material_FL = ""
    DoCmd.SetProperty "Label1176", acPropertyVisible, "-1"
    DoCmd.SetProperty "Description_Material_FL", acPropertyVisible, "-1"
    DoCmd.SetProperty "cmd_Description_Material_FL", acPropertyVisible, "-1"
    Description_Material_01 = ""
    DoCmd.SetProperty "Description_Material_01", acPropertyVisible, "-1"
    DoCmd.SetProperty "cmd_Description_Material_01", acPropertyVisible, "-1"
End Sub

Private Sub cmd_ResetMatDescription_Click()
' Resets and clears all fields and hides all layers of item
' description search when a new category is chosen
    Description_MaterialCategory = ""
    Description_MaterialCategory_01 = ""
    Description_Material = ""
    Description_Material_FL = ""
    DoCmd.SetProperty "Label1176", acPropertyVisible, "0"
    DoCmd.SetProperty "Description_Material_FL", acPropertyVisible, "0"
    DoCmd.SetProperty "cmd_Description_Material_FL", acPropertyVisible, "0"
    Description_Material_01 = ""
    DoCmd.SetProperty "Description_Material_01", acPropertyVisible, "0"
    DoCmd.SetProperty "cmd_Description_Material_01", acPropertyVisible, "0"
    Description_Material_02 = ""
    DoCmd.SetProperty "Description_Material_02", acPropertyVisible, "0"
    DoCmd.SetProperty "cmd_Description_Material_02", acPropertyVisible, "0"
    Description_Material_03 = ""
    DoCmd.SetProperty "Description_Material_03", acPropertyVisible, "0"
    DoCmd.SetProperty "cmd_Description_Material_03", acPropertyVisible, "0"
    Description_MatDeff = ""
End Sub
Private Sub Description_Material_FL_Click()
    Description_MatDeff = "Deffinition: " & [Description_Material_FL].[Column](1) & " (" & [Description_Material_FL].[Column](2) & ")"
End Sub

Private Sub Description_MaterialCategory_01_Click()
' Adds selection to Type/Category from Item/Category and unhides type
    Description_Material_FL = ""
    DoCmd.SetProperty "Label1176", acPropertyVisible, "-1"
    DoCmd.SetProperty "Description_Material_FL", acPropertyVisible, "-1"
    DoCmd.SetProperty "cmd_Description_Material_FL", acPropertyVisible, "-1"
    Description_Material_01 = ""
    DoCmd.SetProperty "Description_Material_01", acPropertyVisible, "-1"
    DoCmd.SetProperty "cmd_Description_Material_01", acPropertyVisible, "-1"
    Description_Material_02 = ""
    DoCmd.SetProperty "Description_Material_02", acPropertyVisible, "0"
    DoCmd.SetProperty "cmd_Description_Material_02", acPropertyVisible, "0"
    Description_Material_03 = ""
    DoCmd.SetProperty "Description_Material_03", acPropertyVisible, "0"
    DoCmd.SetProperty "cmd_Description_Material_03", acPropertyVisible, "0"
    Description_MatDeff = "Deffinition: " & [Description_MaterialCategory_01].[Column](1)

End Sub
Private Sub Description_Material_01_AfterUpdate()
    DoCmd.Requery "Description_Material_01"
    DoCmd.Requery "Description_Material_02"
    DoCmd.Requery "Description_Material_03"
    DoCmd.Requery "Description_MaterialCategory_01"
End Sub

Private Sub Description_Material_01_Click()
' Unhides sub-layers of search Materials
    DoCmd.SetProperty "Description_Material_02", acPropertyVisible, "-1"
    DoCmd.SetProperty "cmd_Description_Material_02", acPropertyVisible, "-1"
    Description_MatDeff = "Deffinition: " & [Description_Material_01].[Column](1)
End Sub

Private Sub Description_Material_02_Click()
' Unhides sub-layers of search Materials
    DoCmd.SetProperty "Description_Material_03", acPropertyVisible, "-1"
    DoCmd.SetProperty "cmd_Description_Material_03", acPropertyVisible, "-1"
    Description_MatDeff = "Deffinition: " & [Description_Material_02].[Column](1)
End Sub
Private Sub Description_Material_03_Click()
    Description_MatDeff = "Deffinition: " & [Description_Material_03].[Column](1)
End Sub
Private Sub Description_MaterialCategory_01_AfterUpdate()
    DoCmd.Requery "Description_Material_01"
    DoCmd.Requery "Description_Material_02"
    DoCmd.Requery "Description_Material_03"
    DoCmd.Requery "Description_MaterialCategory_01"
End Sub
Private Sub Description_Material_02_AfterUpdate()
    DoCmd.Requery "Description_Material_01"
    DoCmd.Requery "Description_Material_02"
    DoCmd.Requery "Description_Material_03"
    DoCmd.Requery "Description_MaterialCategory_01"
End Sub
Private Sub Description_Material_03_AfterUpdate()
    DoCmd.Requery "Description_Material_01"
    DoCmd.Requery "Description_Material_02"
    DoCmd.Requery "Description_Material_03"
    DoCmd.Requery "Description_MaterialCategory_01"
End Sub
Private Sub Description_Material_04_AfterUpdate()
    DoCmd.Requery "Description_Material_01"
    DoCmd.Requery "Description_Material_02"
    DoCmd.Requery "Description_Material_03"
    DoCmd.Requery "Description_MaterialCategory_01"
End Sub
Private Sub Description_Material_05_AfterUpdate()
    DoCmd.Requery "Description_Material_01"
    DoCmd.Requery "Description_Material_02"
    DoCmd.Requery "Description_Material_03"
    DoCmd.Requery "Description_MaterialCategory_01"
'
'  END OF MATERIAL SECTION
'
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
    DoCmd.RunSQL "INSERT INTO bak_ArchiveContents SELECT Record_Type,Text_Caption,File_Name,Text_Created,Location_Grid,Location_Orientation,Location_X,Location_Y,Associated_Find_List,Associated_CommunityFind_List,Associated_Trench_List,Associated_Context_List,Associated_Sample_List,Associated_Building_List,Associated_Room_List,Associated_Masonry_List,Associated_WorkedStone_List,Associated_Survey_List,Associated_Photogrametry_List,Associated_Skeleton_List FROM tbl_ArchiveContents " & _
        "WHERE UID=" & UID
    DoCmd.SetWarnings True
End Sub
