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
    DoCmd.SetProperty "Label1167", acPropertyVisible, "-1"
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

Private Sub Description_Material_FL_Click()
    Description_MatDeff = "Deffinition: " & [Description_Material_FL].[Column](1)
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
    Description_MasonryDeff = "Deffinition: " & [Description_MasonryCategory_01].[Column](1)
End Sub
Private Sub Description_MASONRY_02_Click()
    Description_MasonryDeff = "Deffinition: " & [Description_Masonry_02].[Column](1)
End Sub
Private Sub Description_MASONRY_FL_Click()
    Description_MasonryDeff = "Deffinition: " & [Description_Masonry_FL].[Column](1)
End Sub
Private Sub Description_MASONRY_01_Click()
    Description_MasonryDeff = "Deffinition: " & [Description_Masonry_01].[Column](1)
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

