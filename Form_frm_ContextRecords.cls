VERSION 1.0 CLASS
BEGIN
MultiUse = -1  'True
End
Attribute VB_Name = "Form_frm_ContextRecords"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UpdatePageVisibility()
    Dim cutValues As Variant
    cutValues = Array("Cut", "Structural")
    ' Assume the Tab Control is named "TabControl1"
    ' and "Description_Cut" is at index 0, "Description_Deposit" is at index 1
    If Not IsNull(Me.Context_Type.Value) And InStr(Join(cutValues, ","), Me.Context_Type.Value) > 0 Then
        Me.TabCtl8166.Pages(0).Visible = True
        Me.TabCtl8166.Pages(1).Visible = False
    Else
        Me.TabCtl8166.Pages(0).Visible = False
        Me.TabCtl8166.Pages(1).Visible = True
    End If
End Sub

Private Sub cmd_GoSpecialFinds_Click()
    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "frm_FindsSpecialRecords"

    stLinkCriteria = "[ID_FindID]=" & Me![ListBox_RelatedSpecialFinds]
    DoCmd.OpenForm stDocName, , , stLinkCriteria, , 3
End Sub

Private Sub cmd_GoSamples_Click()
    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "frm_SampleRecords"

    stLinkCriteria = "[ID_SampleID]=" & Me![ListBox_RelatedSamples]
    DoCmd.OpenForm stDocName, , , stLinkCriteria, , 3
End Sub

Private Sub cmd_GoPhotographs_Click()
    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "frm_ArchiveSitePhotos"

    stLinkCriteria = "[UID]=" & Me![ListBox_RelatedPhotos]
    DoCmd.OpenForm stDocName, , , stLinkCriteria, , 3
End Sub

Private Sub cmdFilter_Click()
    Me.Form.Filter = "ID_Context=" & Me.txt_GoTo
    Me.Form.FilterOn = True
End Sub

Private Sub Context_Type_AfterUpdate()
    UpdatePageVisibility()
End Sub
Private Sub Form_Load()
    ' Update the page visibility when the form loads
    UpdatePageVisibility()
End Sub
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


Private Sub cmdRemoveFilter_Click()
    On Error GoTo NoFilter
    DoCmd.RunCommand acCmdRemoveAllFilters
    Me.Form.FilterOn = False
NoFilter:
    MsgBox "You now have no filters set."
End Sub

Private Sub cmdSearch_Click()
    DoCmd.OpenForm "frm_SearchContextRecords", acNormal, "", "", acFormEdit, acNormal
End Sub

Private Sub Form_Current()
    Me.ListBox_RelatedPhotogrammetry.Requery
    Me.ListBox_RelatedPhotos.Requery
    Me.ListBox_RelatedSpecialFinds.Requery
    Me.ListBox_RelatedSamples.Requery
    UpdatePageVisibility
    Me.txtRecs = "Record" & " " & Form.CurrentRecord & " " & "of" & " " & Form.RecordsetClone.RecordCount
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

