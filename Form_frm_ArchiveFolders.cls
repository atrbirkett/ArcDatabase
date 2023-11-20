VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frm_ArchiveFolders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

Private Sub Dating_Period_AfterUpdate()
    On Error GoTo 0
    Dating_MinYear = [Dating_Period].[Column](1)
    Dating_MaxYear = [Dating_Period].[Column](2)
End Sub

Private Sub ID_Building_AfterUpdate()
    On Error GoTo 0
    DoCmd.Requery "Associated_Rooms"
End Sub

Private Sub cmdClose_Click()

    'button to close down the current form
    '(prompting to save changes)
    DoCmd.Close acForm, Me.Name, acSavePrompt
    DoCmd.TransferText acExportDelim, , "lst_Assemblage", CurrentProject.Path & "\Data_Database\lst_Assemblage.csv", True
    
End Sub
