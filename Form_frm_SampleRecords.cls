VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frm_SampleRecords"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmd_GoPhotographs_Click()
    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "frm_ArchiveSitePhotos"
    
    stLinkCriteria = "[UID]=" & Me![ListBox_RelatedPhotos]
    DoCmd.OpenForm stDocName, , , stLinkCriteria, , 3
End Sub

Private Sub cmdRemoveFilter_Click()

      On Error GoTo NoFilter
    
    DoCmd.RunCommand acCmdRemoveAllFilters
    
NoFilter:
    
    MsgBox "You have no filters set."

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
Private Sub cmdSingleSample_Click()
    ' Show the SampleType buttons
    Me.SampleType_Wood.Visible = True
    Me.SampleType_Charcoal.Visible = True
    Me.SampleType_Misc.Visible = True
    Me.Type_SampleType.Visible = True
End Sub
Private Sub cmdBulkSample_Click()
    ' Show the SampleType buttons
    Me.SampleType_Wood.Visible = False
    Me.SampleType_Charcoal.Visible = False
    Me.SampleType_Misc.Visible = False
    Me.Type_SampleType.Visible = False
    Me.TabCtl172.Pages(0).Visible = False
    Me.TabCtl172.Pages(1).Visible = True
    Me.TabCtl172.Pages(3).Visible = False
    Me.TabCtl172.Pages(2).Visible = False
End Sub
Private Sub SampleType_Wood_Click()
    ' Hide other pages and show "Details_Wood"
    Me.TabCtl172.Pages(0).Visible = True
    Me.TabCtl172.Pages(1).Visible = False
    Me.TabCtl172.Pages(3).Visible = False
    Me.TabCtl172.Pages(2).Visible = False
End Sub
Private Sub SampleType_Misc_Click()
    ' Hide other pages and show "Details_Wood"
    Me.TabCtl172.Pages(0).Visible = False
    Me.TabCtl172.Pages(1).Visible = False
    Me.TabCtl172.Pages(3).Visible = True
    Me.TabCtl172.Pages(2).Visible = False
End Sub
Private Sub SampleType_Charcoal_Click()
    ' Hide other pages and show "Details_Wood"
    Me.TabCtl172.Pages(0).Visible = False
    Me.TabCtl172.Pages(1).Visible = False
    Me.TabCtl172.Pages(3).Visible = False
    Me.TabCtl172.Pages(2).Visible = True
End Sub
Private Sub Form_Current()

    Me.txtRecs = "Record" & " " & Form.CurrentRecord & " " & "of" & " " & Form.RecordsetClone.RecordCount

End Sub
Private Sub Form_Load()
    Me.TabCtl172.Pages(0).Visible = False
    Me.TabCtl172.Pages(1).Visible = False
    Me.Details_Charcoal.Visible = False
    Me.TabCtl172.Pages(3).Visible = False
    Me.Details_Misc.Visible = False
    Me.TabCtl172.Pages(2).Visible = False
    Me.Details_Wood.Visible = False
    Me.Type_SampleType.Visible = False
End Sub
