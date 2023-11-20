VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frm_TrenchRecords"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmd_AddContexts_Click()
 DoCmd.OpenForm "frm_ContextRecords", acNormal, "", "", acEdit, acNormal
End Sub
Private Sub txt_GoTo_AfterUpdate()
    Me.Form.Filter = "ID_Trench=" & Me.txt_GoTo
    Me.Form.FilterOn = True
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
    DoCmd.RunSQL "INSERT INTO bak_TrenchRecords SELECT ID_Trench,Location_Grid,Location_Address,Location_X,Location_Y,Date_Start,Date_End,Description_Notes,Mesure_Dimensions,Mesure_TotalArea,Mesure_TotalVolume,Mesure_Maxdepth,Mesure_Mindepth,Mesure_Maxtop,Mesure_Mintop FROM tbl_TrenchRecords " & _
        "WHERE UID=" & UID
    DoCmd.SetWarnings True
End Sub

