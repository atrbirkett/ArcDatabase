VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frm_SkeletonRecords"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmd_AddFinds_Click()
    DoCmd.OpenForm "frm_FindsRecords", acNormal, "", "", acEdit, acNormal
End Sub

Private Sub cmd_AddLevels_Click()
    DoCmd.OpenForm "frm_LevelRecords", acNormal, "", "", acEdit, acNormal
End Sub

Private Sub cmd_AddSamplesBulk_Click()

    DoCmd.OpenForm "frm_SampleBulkRecords", acNormal, "", "", acEdit, acNormal
End Sub

Private Sub cmd_AddSamplesSngl_Click()
    DoCmd.OpenForm "frm_SampleSingleRecords", acNormal, "", "", acEdit, acNormal
End Sub

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
    
NoFilter:
    
    MsgBox "You have no filters set."

End Sub

Private Sub Form_Current()
   
    Me.txtRecs = "Record" & " " & Form.CurrentRecord & " " & "of" & " " & Form.RecordsetClone.RecordCount

End Sub

Private Sub cmdClose_Click()

    'button to close down the current form
    '(prompting to save changes)
    DoCmd.Close acForm, Me.Name, acSavePrompt
    
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


'------------------------------------------------------------
' Form_BeforeUpdate
'
'------------------------------------------------------------
Private Sub Form_BeforeUpdate(Cancel As Integer)
    DoCmd.SetWarnings False
    DoCmd.RunSQL "INSERT INTO tbl_SkeletonRecords_bak SELECT * FROM tbl_SkeletonRecords " & _
        "WHERE UID=" & UID
    DoCmd.SetWarnings True
End Sub


