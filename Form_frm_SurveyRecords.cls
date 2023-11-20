VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frm_SurveyRecords"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub txt_GoTo_AfterUpdate()
    Me.Form.Filter = "UID=" & Me.txt_GoTo
    Me.Form.FilterOn = True
End Sub
Private Sub cmd_Add_Dating_Period_Click()
     If IsNull(Dating_Period) Then Dating_Period = Dating_Period_List Else: Dating_Period = Dating_Period & "; " & Dating_Period_List
End Sub

Private Sub cmdRemoveFilter_Click()
      On Error GoTo NoFilter
    DoCmd.RunCommand acCmdRemoveAllFilters
    Me.Form.FilterOn = False
NoFilter:
    MsgBox "You now have no filters set."
End Sub
'
'  BEGINING OF CLASSIFICATION SECTION
'
Private Sub cmd_Description_Classification_01_Click()
' Adds selection to Classification/Category from Classification/Category
     Description_Classification = Description_Classification & "; " & Description_Classification_01
End Sub
Private Sub cmdSearch_Click()
 DoCmd.OpenForm "frm_SearchSurveyRecords", acNormal, "", "", acFormEdit, acNormal
End Sub
Private Sub cmd_Description_Classification_02_Click()
' Adds selection to Classification/Category from Classification/Category
     Description_Classification = Description_Classification & "; " & Description_Classification_02
End Sub

Private Sub cmd_Description_Classification_03_Click()
' Adds selection to Classification/Category from Classification/Category
     Description_Classification = Description_Classification & "; " & Description_Classification_03
End Sub

Private Sub cmd_Description_Classification_04_Click()
' Adds selection to Classification/Category from Classification/Category
     Description_Classification = Description_Classification & "; " & Description_Classification_04
End Sub

Private Sub cmd_Description_Classification_05_Click()
' Adds selection to Classification/Category from Classification/Category
     Description_Classification = Description_Classification & "; " & Description_Classification_05
End Sub

Private Sub cmd_Description_Classification_FL_Click()
' Adds selection to Type/Category from Item/Category
    Description_Classification = Description_Classification & "; " & Description_Classification_FL
End Sub

Private Sub cmd_Description_ClassificationCategory_01_Click()
' Adds selection to Type/Category from Item/Category and unhides type
    Description_Classification_FL = ""
    DoCmd.SetProperty "Description_Classification_FL", acPropertyVisible, "-1"
    DoCmd.SetProperty "cmd_Description_Classification_FL", acPropertyVisible, "-1"
    Description_Classification_01 = ""
    DoCmd.SetProperty "Description_Classification_01", acPropertyVisible, "-1"
    DoCmd.SetProperty "cmd_Description_Classification_01", acPropertyVisible, "-1"
    Description_Classification = Description_MaterialClassClass & "" & Description_ClassificationCategory_01

End Sub
Private Sub cmd_Rest_Classification_Click()
' Resets and clears all fields and hides all layers of item
' description search when a new category is chosen
    Description_ClassificationCategory = ""
    Description_ClassificationCategory_01 = ""
    Description_Classification = ""
    Description_Classification_FL = ""
    DoCmd.SetProperty "Description_Classification_FL", acPropertyVisible, "0"
    DoCmd.SetProperty "cmd_Description_Classification_FL", acPropertyVisible, "0"
    Description_Classification_01 = ""
    DoCmd.SetProperty "Description_Classification_01", acPropertyVisible, "0"
    DoCmd.SetProperty "cmd_Description_Classification_01", acPropertyVisible, "0"
    Description_Classification_02 = ""
    DoCmd.SetProperty "Description_Classification_02", acPropertyVisible, "0"
    DoCmd.SetProperty "cmd_Description_Classification_02", acPropertyVisible, "0"
    Description_Classification_03 = ""
    DoCmd.SetProperty "Description_Classification_03", acPropertyVisible, "0"
    DoCmd.SetProperty "cmd_Description_Classification_03", acPropertyVisible, "0"
    Description_Classification_04 = ""
    DoCmd.SetProperty "Description_Classification_04", acPropertyVisible, "0"
    DoCmd.SetProperty "cmd_Description_Classification_04", acPropertyVisible, "0"
    Description_Classification_05 = ""
    DoCmd.SetProperty "Description_Classification_05", acPropertyVisible, "0"
    DoCmd.SetProperty "cmd_Description_Classification_05", acPropertyVisible, "0"
    Description_ClassificationDeff = ""

End Sub

Private Sub Description_Classification_01_AfterUpdate()
    DoCmd.Requery "Description_Classification_01"
    DoCmd.Requery "Description_Classification_02"
    DoCmd.Requery "Description_Classification_03"
    DoCmd.Requery "Description_ClassificationCategory_01"
End Sub

Private Sub Description_ClassificationCategory_01_Click()
' Adds selection to Type/Category from Item/Category and unhides type
    Description_Classification_01 = ""
    Description_Classification_02 = ""
    DoCmd.SetProperty "Description_Classification_02", acPropertyVisible, "0"
    DoCmd.SetProperty "cmd_Description_Classification_02", acPropertyVisible, "0"
    Description_Classification_03 = ""
    DoCmd.SetProperty "Description_Classification_03", acPropertyVisible, "0"
    DoCmd.SetProperty "cmd_Description_Classification_03", acPropertyVisible, "0"
    Description_ClassificationDeff = "Definition: " & [Description_ClassificationCategory_01].[Column](1)
End Sub

Private Sub Description_Classification_01_Click()
' Unhides sub-layers of search Classifications
    DoCmd.SetProperty "Description_Classification_02", acPropertyVisible, "-1"
    DoCmd.SetProperty "cmd_Description_Classification_02", acPropertyVisible, "-1"
    Description_ClassificationDeff = "Definition: " & [Description_Classification_01].[Column](1)
End Sub

Private Sub Description_Classification_02_Click()
' Unhides sub-layers of search Classifications
    DoCmd.SetProperty "Description_Classification_03", acPropertyVisible, "-1"
    DoCmd.SetProperty "cmd_Description_Classification_03", acPropertyVisible, "-1"
    Description_ClassificationDeff = "Definition: " & [Description_Classification_02].[Column](1)
End Sub

Private Sub Description_Classification_03_Click()
' Unhides sub-layers of search Classifications
    DoCmd.SetProperty "Description_Classification_04", acPropertyVisible, "-1"
    DoCmd.SetProperty "cmd_Description_Classification_04", acPropertyVisible, "-1"
    Description_ClassificationDeff = "Definition: " & [Description_Classification_03].[Column](1)
End Sub

Private Sub Description_Classification_04_Click()
' Unhides sub-layers of search Classifications
    DoCmd.SetProperty "Description_Classification_05", acPropertyVisible, "-1"
    DoCmd.SetProperty "cmd_Description_Classification_05", acPropertyVisible, "-1"
    Description_ClassificationDeff = "Definition: " & [Description_Classification_04].[Column](1)
End Sub
Private Sub Description_Classification_FL_Click()
    Description_ClassificationDeff = "Definition: " & [Description_Classification_FL].[Column](1)
End Sub
Private Sub Description_ClassificationCategory_01_AfterUpdate()
    DoCmd.Requery "Description_Classification_01"
    DoCmd.Requery "Description_Classification_02"
    DoCmd.Requery "Description_Classification_03"
    DoCmd.Requery "Description_ClassificationCategory_01"
End Sub
Private Sub Description_Classification_02_AfterUpdate()
    DoCmd.Requery "Description_Classification_01"
    DoCmd.Requery "Description_Classification_02"
    DoCmd.Requery "Description_Classification_03"
    DoCmd.Requery "Description_ClassificationCategory_01"
End Sub
Private Sub Description_Classification_03_AfterUpdate()

    DoCmd.Requery "Description_Classification_01"
    DoCmd.Requery "Description_Classification_02"
    DoCmd.Requery "Description_Classification_03"
    DoCmd.Requery "Description_ClassificationCategory_01"
End Sub
Private Sub Description_Classification_04_AfterUpdate()
    DoCmd.Requery "Description_Classification_01"
    DoCmd.Requery "Description_Classification_02"
    DoCmd.Requery "Description_Classification_03"
    DoCmd.Requery "Description_ClassificationCategory_01"
End Sub
Private Sub Description_Classification_05_AfterUpdate()
    DoCmd.Requery "Description_Classification_01"
    DoCmd.Requery "Description_Classification_02"
    DoCmd.Requery "Description_Classification_03"
    DoCmd.Requery "Description_ClassificationCategory_01"
'
'  END OF CLASSIFICATION SECTION
'
End Sub


Private Sub Dating_Period_List_DblClick(Cancel As Integer)
     If IsNull(Dating_Period) Then Dating_Period = Dating_Period_List Else: Dating_Period = Dating_Period & "; " & Dating_Period_List
End Sub

Private Sub Form_Current()
   
    Me.txtRecs = "Record" & " " & Form.CurrentRecord & " " & "of" & " " & Form.RecordsetClone.RecordCount

End Sub


Private Sub Form_BeforeUpdate(Cancel As Integer)
    DoCmd.SetWarnings False
    DoCmd.RunSQL "INSERT INTO bak_Surveynotes SELECT ID_Name,Location_Grid,Location_Address,Location_X,Location_Y,Location_Coverage,Description_EventType,Description_Findings,Description_Notes,Dating_Period,SurveyDate_Start,SurveyDate_End,Created_Date,Associated_Find_List,Associated_Trench_List,Associated_Context_List,Associated_Skeleton_List,Associated_Sample_List,Associated_Building_List,Associated_Masonry_List,Associated_WorkedStone_List,Associated_Photogrammetry_List,Associated_Room_List,Associated_Survey_List FROM tbl_SurveyNotes " & _
        "WHERE UID=" & UID
    DoCmd.SetWarnings True
End Sub
Private Sub cmd_Classification_TypeCategory_01_Click()
' Adds selection to Type/Category from Item/Category and unhides Type
    If IsNull(Classification_Type) Then Classification_Type = Classification_TypeCategory_01 Else: Classification_Type = Classification_Type & "; " & Classification_TypeCategory_01
    DoCmd.Requery "Classification_Type_01"
    
    Classification_Type_FL = ""
    DoCmd.SetProperty "Classification_Type_FL", acPropertyVisible, "-1"
    DoCmd.SetProperty "cmd_Classification_Type_FL", acPropertyVisible, "-1"
    Classification_Type_01 = ""
    DoCmd.SetProperty "Classification_Type_01", acPropertyVisible, "-1"
    DoCmd.SetProperty "cmd_Classification_Type_01", acPropertyVisible, "-1"
    Classification_Type = Classification_Type & "" & Classification_TypeCategory_01

End Sub
Private Sub cmd_Rest_Type_Click()
' Resets and clears all fields and hides all layers of item
' Classification search when a new category is chosen
    Classification_TypeCategory = ""
    Classification_TypeCategory_01 = ""
    Classification_Type = ""
    Classification_Type_FL = ""
    DoCmd.SetProperty "Classification_Type_FL", acPropertyVisible, "0"
    DoCmd.SetProperty "cmd_Classification_Type_FL", acPropertyVisible, "0"
    Classification_Type_01 = ""
    DoCmd.SetProperty "Classification_Type_01", acPropertyVisible, "0"
    DoCmd.SetProperty "cmd_Classification_Type_01", acPropertyVisible, "0"
    Classification_Type_02 = ""
    DoCmd.SetProperty "Classification_Type_02", acPropertyVisible, "FALSE"
    DoCmd.SetProperty "cmd_Classification_Type_02", acPropertyVisible, "FALSE"
    Classification_Type_03 = ""
    DoCmd.SetProperty "Classification_Type_03", acPropertyVisible, "0"
    DoCmd.SetProperty "cmd_Classification_Type_03", acPropertyVisible, "0"
    Classification_TypeDeff = ""

End Sub
Private Sub cmd_Classification_Type_01_Click()
' Adds selection to Type/Category from Type/Category
    If IsNull(Classification_Type) Then Classification_Type = Classification_Type_01 Else: Classification_Type = Classification_Type & "; " & Classification_Type_01
    DoCmd.Requery "Classification_Type_02"
End Sub

Private Sub cmd_Classification_Type_02_Click()
' Adds selection to Type/Category from Type/Category
     Classification_Type = Classification_Type & "; " & Classification_Type_02
End Sub

Private Sub cmd_Classification_Type_03_Click()
' Adds selection to Type/Category from Type/Category
     Classification_Type = Classification_Type & "; " & Classification_Type_03
End Sub


Private Sub cmd_Classification_Type_FL_Click()
' Adds selection to Type/Category from Item/Category
     If IsNull(Classification_Type) Then Classification_Type = Classification_Type_FL Else: Classification_Type = Classification_Type & "; " & Classification_Type_FL
End Sub

Private Sub Classification_Type_01_AfterUpdate()
    DoCmd.Requery "Classification_Type_01"
    DoCmd.Requery "Classification_Type_02"
    DoCmd.Requery "Classification_Type_03"
    DoCmd.Requery "Classification_TypeCategory_01"
End Sub

Private Sub Classification_TypeCategory_01_Click()
' Adds selection to Type/Category from Item/Category and unhides Type
    Classification_Type_01 = ""
    Classification_Type_02 = ""
    DoCmd.SetProperty "Classification_Type_02", acPropertyVisible, "0"
    DoCmd.SetProperty "cmd_Classification_Type_02", acPropertyVisible, "0"
    Classification_Type_03 = ""
    DoCmd.SetProperty "Classification_Type_03", acPropertyVisible, "0"
    DoCmd.SetProperty "cmd_Classification_Type_03", acPropertyVisible, "0"
    Classification_Type_04 = ""
    Classification_TypeDeff = "Definition: " & [Classification_TypeCategory_01].[Column](1)
End Sub

Private Sub Classification_Type_01_Click()
' Unhides sub-layers of search Types
    DoCmd.SetProperty "Classification_Type_02", acPropertyVisible, "-1"
    DoCmd.SetProperty "cmd_Classification_Type_02", acPropertyVisible, "-1"
    Classification_TypeDeff = "Definition: " & [Classification_Type_01].[Column](1)
End Sub

Private Sub Classification_Type_02_Click()
' Unhides sub-layers of search Types
    DoCmd.SetProperty "Classification_Type_03", acPropertyVisible, "-1"
    DoCmd.SetProperty "cmd_Classification_Type_03", acPropertyVisible, "-1"
    Classification_TypeDeff = "Definition: " & [Classification_Type_02].[Column](1)
End Sub

Private Sub Classification_Type_03_Click()
' Unhides sub-layers of search Types
    Classification_TypeDeff = "Definition: " & [Classification_Type_03].[Column](1)
End Sub

Private Sub Classification_Type_04_Click()
' Unhides sub-layers of search Types
    DoCmd.SetProperty "Classification_Type_05", acPropertyVisible, "-1"
    DoCmd.SetProperty "cmd_Classification_Type_05", acPropertyVisible, "-1"
    Classification_TypeDeff = "Definition: " & [Classification_Type_04].[Column](1)
End Sub
Private Sub Classification_Type_FL_Click()
    Classification_TypeDeff = "Definition: " & [Classification_Type_FL].[Column](1) & " (" & [Classification_Type_FL].[Column](2) & ")"
End Sub
Private Sub Classification_TypeCategory_01_AfterUpdate()
    DoCmd.Requery "Classification_Type_01"
    DoCmd.Requery "Classification_Type_02"
    DoCmd.Requery "Classification_Type_03"
    DoCmd.Requery "Classification_TypeCategory_01"
End Sub
Private Sub Classification_Type_02_AfterUpdate()
    DoCmd.Requery "Classification_Type_01"
    DoCmd.Requery "Classification_Type_02"
    DoCmd.Requery "Classification_Type_03"
    DoCmd.Requery "Classification_TypeCategory_01"
End Sub
Private Sub Classification_Type_03_AfterUpdate()

    DoCmd.Requery "Classification_Type_01"
    DoCmd.Requery "Classification_Type_02"
    DoCmd.Requery "Classification_Type_03"
    DoCmd.Requery "Classification_TypeCategory_01"
'
'  END OF Type SECTION
'
End Sub
Private Sub cmd_Classification_EventCategory_01_Click()
' Adds selection to Event/Category from Item/Category and unhides Event
     If IsNull(Classification_Event) Then Classification_Event = Classification_EventCategory_01 Else: Classification_Event = Classification_Event & "; " & Classification_EventCategory_01
     DoCmd.Requery "Classification_Event_01"
    Classification_Event_FL = ""
    DoCmd.SetProperty "Classification_Event_FL", acPropertyVisible, "-1"
    DoCmd.SetProperty "cmd_Classification_Event_FL", acPropertyVisible, "-1"
    Classification_Event_01 = ""
    DoCmd.SetProperty "Classification_Event_01", acPropertyVisible, "-1"
    DoCmd.SetProperty "cmd_Classification_Event_01", acPropertyVisible, "-1"
End Sub
Private Sub cmd_Rest_Event_Click()
' Resets and clears all fields and hides all layers of item
' Classification search when a new category is chosen
    Classification_EventCategory = ""
    Classification_EventCategory_01 = ""
    Classification_Event = ""
    Classification_Event_FL = ""
    DoCmd.SetProperty "Classification_Event_FL", acPropertyVisible, "0"
    DoCmd.SetProperty "cmd_Classification_Event_FL", acPropertyVisible, "0"
    Classification_Event_01 = ""
    DoCmd.SetProperty "Classification_Event_01", acPropertyVisible, "0"
    DoCmd.SetProperty "cmd_Classification_Event_01", acPropertyVisible, "0"
    Classification_Event_02 = ""
    DoCmd.SetProperty "Classification_Event_02", acPropertyVisible, "FALSE"
    DoCmd.SetProperty "cmd_Classification_Event_02", acPropertyVisible, "FALSE"
    Classification_Event_03 = ""
    DoCmd.SetProperty "Classification_Event_03", acPropertyVisible, "0"
    DoCmd.SetProperty "cmd_Classification_Event_03", acPropertyVisible, "0"
    Classification_EventDeff = ""

End Sub
Private Sub cmd_Classification_Event_01_Click()
' Adds selection to Event/Category from Event/Category
     If IsNull(Classification_Event) Then Classification_Event = Classification_Event_01 Else: Classification_Event = Classification_Event & "; " & Classification_Event_01
     DoCmd.Requery "Classification_Event_02"
End Sub

Private Sub cmd_Classification_Event_02_Click()
' Adds selection to Event/Category from Event/Category
     Classification_Event = Classification_Event & "; " & Classification_Event_02
End Sub

Private Sub cmd_Classification_Event_03_Click()
' Adds selection to Event/Category from Event/Category
     Classification_Event = Classification_Event & "; " & Classification_Event_03
End Sub


Private Sub cmd_Classification_Event_FL_Click()
' Adds selection to Event/Category from Item/Category
    If IsNull(Classification_Event) Then Classification_Event = Classification_Event_FL Else: Classification_Event = Classification_Event & "; " & Classification_Event_FL
End Sub

Private Sub Classification_Event_01_AfterUpdate()
    DoCmd.Requery "Classification_Event_01"
    DoCmd.Requery "Classification_Event_02"
    DoCmd.Requery "Classification_Event_03"
    DoCmd.Requery "Classification_EventCategory_01"
End Sub

Private Sub Classification_EventCategory_01_Click()
' Adds selection to Event/Category from Item/Category and unhides Event
    Classification_Event_01 = ""
    Classification_Event_02 = ""
    DoCmd.SetProperty "Classification_Event_02", acPropertyVisible, "0"
    DoCmd.SetProperty "cmd_Classification_Event_02", acPropertyVisible, "0"
    Classification_Event_03 = ""
    DoCmd.SetProperty "Classification_Event_03", acPropertyVisible, "0"
    DoCmd.SetProperty "cmd_Classification_Event_03", acPropertyVisible, "0"
    Classification_Event_04 = ""
        Classification_Event_FL = ""
    Classification_EventDeff = "Definition: " & [Classification_EventCategory_01].[Column](1)
End Sub

Private Sub Classification_Event_01_Click()
' Unhides sub-layers of search Events
    DoCmd.SetProperty "Classification_Event_02", acPropertyVisible, "-1"
    DoCmd.SetProperty "cmd_Classification_Event_02", acPropertyVisible, "-1"
    Classification_EventDeff = "Definition: " & [Classification_Event_01].[Column](1)
End Sub

Private Sub Classification_Event_02_Click()
' Unhides sub-layers of search Events
    DoCmd.SetProperty "Classification_Event_03", acPropertyVisible, "-1"
    DoCmd.SetProperty "cmd_Classification_Event_03", acPropertyVisible, "-1"
    Classification_EventDeff = "Definition: " & [Classification_Event_02].[Column](1)
End Sub

Private Sub Classification_Event_03_Click()
' Unhides sub-layers of search Events
    Classification_EventDeff = "Definition: " & [Classification_Event_03].[Column](1)
End Sub

Private Sub Classification_Event_04_Click()
' Unhides sub-layers of search Events
    DoCmd.SetProperty "Classification_Event_05", acPropertyVisible, "-1"
    DoCmd.SetProperty "cmd_Classification_Event_05", acPropertyVisible, "-1"
    Classification_EventDeff = "Definition: " & [Classification_Event_04].[Column](1)
End Sub
Private Sub Classification_Event_FL_Click()
    Classification_EventDeff = "Definition: " & [Classification_Event_FL].[Column](1) & " (" & [Classification_Event_FL].[Column](2) & ")"
End Sub
Private Sub Classification_EventCategory_01_AfterUpdate()
    DoCmd.Requery "Classification_Event_01"
    DoCmd.Requery "Classification_Event_02"
    DoCmd.Requery "Classification_Event_03"
    DoCmd.Requery "Classification_EventCategory_01"
End Sub
Private Sub Classification_Event_02_AfterUpdate()
    DoCmd.Requery "Classification_Event_01"
    DoCmd.Requery "Classification_Event_02"
    DoCmd.Requery "Classification_Event_03"
    DoCmd.Requery "Classification_EventCategory_01"
End Sub
Private Sub Classification_Event_03_AfterUpdate()

    DoCmd.Requery "Classification_Event_01"
    DoCmd.Requery "Classification_Event_02"
    DoCmd.Requery "Classification_Event_03"
    DoCmd.Requery "Classification_EventCategory_01"
End Sub

'
'  END OF Event SECTION
