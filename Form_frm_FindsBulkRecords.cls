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
    Me.Description_Features_01.Visible = False
    Me.cmd_Description_Material_01.Visible = False
    Me.Description_Features_02.Visible = False
    Me.cmd_Description_Material_02.Visible = False
    Me.Description_Features_03.Visible = False
    Me.cmd_Description_Material_03.Visible = False
    Me.Description_Features_04.Visible = False
    Me.cmd_Description_Features_04.Visible = False
    Me.Description_Features_05.Visible = False
    Me.cmd_Description_Features_05.Visible = False
    Me.Description_Material_01.Visible = False
    Me.cmd_Description_Features_01.Visible = False
    Me.Description_Material_02.Visible = False
    Me.cmd_Description_Features_02.Visible = False
    Me.Description_Material_03.Visible = False
    Me.cmd_Description_Features_03.Visible = False
    
End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)
    DoCmd.SetWarnings False
    DoCmd.RunSQL "INSERT INTO bak_FindsBulkRecords SELECT ID_BagID,ID_AssemblageID,ID_Unstratified,ID_Trench,ID_Context,Location_Grid,Count_NumberofType,Count_CountIsEstimation,Count_Weight,Material_Classification,Material_Type,Find_Classificiation,Find_Type,Find_SpecificName,Dating_Period,Dating_MinYear,Dating_MaxYear,Description_Form,Description_Colour,Description_Decoration,Description_Interp,Description_Notes FROM tbl_FindsBulkRecords " & _
        "WHERE UID=" & UID
    DoCmd.SetWarnings True
    
End Sub

Private Sub Form_Load()
    Me.Description_Features_01.Visible = False
    Me.cmd_Description_Material_01.Visible = False
    Me.Description_Features_02.Visible = False
    Me.cmd_Description_Material_02.Visible = False
    Me.Description_Features_03.Visible = False
    Me.cmd_Description_Material_03.Visible = False
    Me.Description_Features_04.Visible = False
    Me.cmd_Description_Features_04.Visible = False
    Me.Description_Features_05.Visible = False
    Me.cmd_Description_Features_05.Visible = False
    Me.Description_Material_01.Visible = False
    Me.cmd_Description_Features_01.Visible = False
    Me.Description_Material_02.Visible = False
    Me.cmd_Description_Features_02.Visible = False
    Me.Description_Material_03.Visible = False
    Me.cmd_Description_Features_03.Visible = False
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
'  BEGINING OF MATERIAL SECTION
'
Private Sub cmd_Description_Material_01_Click()
' Adds selection to Material/Category from Material/Category
          If IsNull(Description_Material) Then Description_Material = Description_Material_01 Else: Description_Material = Description_Material & "; " & Description_Material_01
     DoCmd.Requery "Description_Material"
End Sub

Private Sub cmd_Description_Material_02_Click()
' Adds selection to Material/Category from Material/Category
          If IsNull(Description_Material) Then Description_Material = Description_Material_02 Else: Description_Material = Description_Material & "; " & Description_Material_02
     DoCmd.Requery "Description_Material"
End Sub

Private Sub cmd_Description_Material_03_Click()
' Adds selection to Material/Category from Material/Category
     If IsNull(Description_Material) Then Description_Material = Description_Material_03 Else: Description_Material = Description_Material & "; " & Description_Material_03
     DoCmd.Requery "Description_Material"
End Sub

Private Sub cmd_Description_MatCategory_01_Click()
    If IsNull(Description_MaterialCategory) Then Description_MaterialCategory = Description_MaterialCategory_01 Else: Description_MaterialCategory = Description_MaterialCategory & "; " & Description_MaterialCategory_01
End Sub
Private Sub Description_MaterialCategory_01_DblClick(Cancel As Integer)
    If IsNull(Description_MaterialCategory) Then Description_MaterialCategory = Description_MaterialCategory_01 Else: Description_MaterialCategory = Description_MaterialCategory & "; " & Description_MaterialCategory_01
End Sub
Private Sub cmd_Description_Material_FL_Click()
' Adds selection to Type/Category from Item/Category
    If IsNull(Description_Material) Then Description_Material = Description_Material_FL Else: Description_Material = Description_Material & "; " & Description_Material_FL
    DoCmd.Requery "Description_Material"
End Sub
Private Sub Description_Material_FL_DblClick(Cancel As Integer)
' Adds selection to Type/Category from Item/Category
    If IsNull(Description_Material) Then Description_Material = Description_Material_FL Else: Description_Material = Description_Material & "; " & Description_Material_FL
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
'  BEGINING OF FEATURES SECTION
'
Private Sub cmd_Description_FeaturesCategory_01_Click()
      If IsNull(Description_FeaturesCategory) Then Description_FeaturesCategory = Description_FeaturesCategory_01 Else: Description_FeaturesCategory = Description_FeaturesCategory & "; " & Description_FeaturesCategory_01
End Sub
Private Sub Description_FeaturesCategory_01_DblClick(Cancel As Integer)
      If IsNull(Description_FeaturesCategory) Then Description_FeaturesCategory = Description_FeaturesCategory_01 Else: Description_FeaturesCategory = Description_FeaturesCategory & "; " & Description_FeaturesCategory_01
End Sub
Private Sub Description_Features_01_DblClick(Cancel As Integer)
' Adds selection to Features/Category from Features/Category
      If IsNull(Description_Features) Then Description_Features = Description_Features_01 Else: Description_Features = Description_Features & "; " & Description_Features_01
End Sub
Private Sub cmd_Description_Features_01_Click()
' Adds selection to Features/Category from Features/Category
      If IsNull(Description_Features) Then Description_Features = Description_Features_01 Else: Description_Features = Description_Features & "; " & Description_Features_01
End Sub
Private Sub Description_Features_02_DblClick(Cancel As Integer)
' Adds selection to Features/Category from Features/Category
     Description_Features = Description_Features & "; " & Description_Features_02
End Sub
Private Sub cmd_Description_Features_02_Click()
' Adds selection to Features/Category from Features/Category
     Description_Features = Description_Features & "; " & Description_Features_02
End Sub
Private Sub Description_Features_03_DblClick(Cancel As Integer)
' Adds selection to Features/Category from Features/Category
     Description_Features = Description_Features & "; " & Description_Features_03
End Sub
Private Sub cmd_Description_Features_03_Click()
' Adds selection to Features/Category from Features/Category
     Description_Features = Description_Features & "; " & Description_Features_03
End Sub
Private Sub cmd_Description_Features_04_Click()
' Adds selection to Features/Category from Features/Category
     Description_Features = Description_Features & "; " & Description_Features_04
End Sub
Private Sub Description_Features_DblClick(Cancel As Integer)
' Adds selection to Features/Category from Features/Category
     Description_Features = Description_Features & "; " & Description_Features_04
End Sub
Private Sub cmd_Description_Features_05_Click()
' Adds selection to Features/Category from Features/Category
     Description_Features = Description_Features & "; " & Description_Features_05
End Sub
Private Sub Description_Features_05_DblClick(Cancel As Integer)
' Adds selection to Features/Category from Features/Category
     Description_Features = Description_Features & "; " & Description_Features_05
End Sub
Private Sub cmd_Description_Features_FL_Click()
' Adds selection to Type/Category from Item/Category
  If IsNull(Description_FeaturesCategory) Then Description_Features = Description_Features_FL Else: Description_Features = Description_Features & "; " & Description_Features_FL
End Sub
Private Sub Description_Features_FL_DblClick(Cancel As Integer)
' Adds selection to Type/Category from Item/Category
  If IsNull(Description_FeaturesCategory) Then Description_Features = Description_Features_FL Else: Description_Features = Description_Features & "; " & Description_Features_FL
End Sub

Private Sub cmd_Reset_FeaturesDescription_Click()
' Resets and clears all fields and hides all layers of item
' description search when a new category is chosen
    Description_FeaturesCategory_01 = ""
    Description_FeaturesCategory = ""
    Description_Features = ""
    Description_Features_FL = ""
    Description_Features_01 = ""
    Description_FeaturesDeff = ""
    DoCmd.SetProperty "Description_Features_01", acPropertyVisible, "0"
    DoCmd.SetProperty "cmd_Description_Features_01", acPropertyVisible, "0"
    Description_Features_02 = ""
    DoCmd.SetProperty "Description_Features_02", acPropertyVisible, "0"
    DoCmd.SetProperty "cmd_Description_Features_02", acPropertyVisible, "0"
    Description_Features_03 = ""
    DoCmd.SetProperty "Description_Features_03", acPropertyVisible, "0"
    DoCmd.SetProperty "cmd_Description_Features_03", acPropertyVisible, "0"
        Description_Features_04 = ""
    DoCmd.SetProperty "Description_Features_04", acPropertyVisible, "0"
    DoCmd.SetProperty "cmd_Description_Features_04", acPropertyVisible, "0"
        Description_Features_05 = ""
    DoCmd.SetProperty "Description_Features_05", acPropertyVisible, "0"
    DoCmd.SetProperty "cmd_Description_Features_05", acPropertyVisible, "0"
    Description_FeaturesDeff = ""
End Sub
Private Sub Description_Features_01_AfterUpdate()
    DoCmd.Requery "Description_Features_01"
    DoCmd.Requery "Description_Features_02"
    DoCmd.Requery "Description_Features_03"
    DoCmd.Requery "Description_Features_05"
    DoCmd.Requery "Description_Features_04"
    DoCmd.Requery "Description_FeaturesCategory_01"
End Sub

Private Sub Description_Features_02_AfterUpdate()
    DoCmd.Requery "Description_Features_01"
    DoCmd.Requery "Description_Features_02"
    DoCmd.Requery "Description_Features_03"
        DoCmd.Requery "Description_Features_05"
    DoCmd.Requery "Description_Features_04"
    DoCmd.Requery "Description_FeaturesCategory_01"
End Sub
Private Sub Description_Features_03_AfterUpdate()
    DoCmd.Requery "Description_Features_01"
    DoCmd.Requery "Description_Features_02"
    DoCmd.Requery "Description_Features_03"
        DoCmd.Requery "Description_Features_05"
    DoCmd.Requery "Description_Features_04"
    DoCmd.Requery "Description_FeaturesCategory_01"
End Sub
Private Sub Description_FeaturesCategory_01_AfterUpdate()
    DoCmd.Requery "Description_Features_01"
    DoCmd.Requery "Description_Features_02"
    DoCmd.Requery "Description_Features_03"
        DoCmd.Requery "Description_Features_05"
    DoCmd.Requery "Description_Features_04"
    DoCmd.Requery "Description_FeaturesCategory_01"
End Sub
Private Sub Description_FeaturesCategory_01_Click()
' Adds selection to Type/Category from Item/Category and unhides type
    Description_Features_02 = ""
    DoCmd.SetProperty "Description_Features_02", acPropertyVisible, "0"
    DoCmd.SetProperty "cmd_Description_Features_02", acPropertyVisible, "0"
    Description_Features_03 = ""
    DoCmd.SetProperty "Description_Features_03", acPropertyVisible, "0"
    DoCmd.SetProperty "cmd_Description_Features_03", acPropertyVisible, "0"
        Description_Features_04 = ""
    DoCmd.SetProperty "Description_Features_04", acPropertyVisible, "0"
    DoCmd.SetProperty "cmd_Description_Features_04", acPropertyVisible, "0"
            Description_Features_05 = ""
    DoCmd.SetProperty "Description_Features_05", acPropertyVisible, "0"
    DoCmd.SetProperty "cmd_Description_Features_05", acPropertyVisible, "0"
    ' Adds selection to Type/Category from Item/Category and unhides type
    Description_Features_FL = ""
    DoCmd.SetProperty "Description_Features_FL", acPropertyVisible, "-1"
    DoCmd.SetProperty "cmd_Description_Features_FL", acPropertyVisible, "-1"
    Description_Features_01 = ""
    DoCmd.SetProperty "Description_Features_01", acPropertyVisible, "-1"
    DoCmd.SetProperty "cmd_Description_Features_01", acPropertyVisible, "-1"
    Description_FeaturesDeff = "Deffinition: " & [Description_FeaturesCategory_01].[Column](1)
End Sub
Private Sub Description_Features_01_Click()
' Unhides sub-layers of search Featuress
    DoCmd.SetProperty "Description_Features_02", acPropertyVisible, "-1"
    DoCmd.SetProperty "cmd_Description_Features_02", acPropertyVisible, "-1"
    Description_FeaturesDeff = "Deffinition: " & [Description_Features_01].[Column](1)
End Sub
Private Sub Description_Features_02_Click()
' Unhides sub-layers of search Featuress
    DoCmd.SetProperty "Description_Features_03", acPropertyVisible, "-1"
    DoCmd.SetProperty "cmd_Description_Features_03", acPropertyVisible, "-1"
    Description_FeaturesDeff = "Deffinition: " & [Description_Features_02].[Column](1)
End Sub
Private Sub Description_Features_03_Click()
' Unhides sub-layers of search Featuress
    DoCmd.SetProperty "Description_Features_04", acPropertyVisible, "-1"
    DoCmd.SetProperty "cmd_Description_Features_04", acPropertyVisible, "-1"
    Description_FeaturesDeff = "Deffinition: " & [Description_Features_03].[Column](1)
End Sub
Private Sub Description_Features_04_Click()
' Unhides sub-layers of search Featuress
    DoCmd.SetProperty "Description_Features_05", acPropertyVisible, "-1"
    DoCmd.SetProperty "cmd_Description_Features_05", acPropertyVisible, "-1"
    Description_FeaturesDeff = "Deffinition: " & [Description_Features_04].[Column](1)
End Sub
Private Sub Description_Features_05_Click()
' Unhides sub-layers of search Featuress
    Description_FeaturesDeff = "Deffinition: " & [Description_Features_05].[Column](1)
End Sub
Private Sub Description_Features_FL_Click()
      Description_FeaturesDeff = "Deffinition: " & [Description_Features_FL].[Column](1) & " (" & [Description_Features_FL].[Column](2) & ")"
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
