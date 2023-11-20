VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frm_PhotogrammetryRecords"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Private Sub cmdBackup_Click()
    On Error GoTo cmdBackup_Click_Err
    ' Backup Button (Form)
    DoCmd.RunCommand acCmdExportXML
    MsgBox "Backup Made Successfully"
cmdBackup_Click_Exit:
    Exit Sub
cmdBackup_Click_Err:
    MsgBox Error$
    Resume cmdBackup_Click_Exit
End Sub

Private Sub cmdAddToList_Click()
    Associated_Find_List = Associated_Find
    Associated_Survey_List = Associated_Survey
    Associated_Building_List = Associated_Building
    Associated_Masonry_List = Associated_Masonry
    Associated_WorkedStone_List = Associated_WorkedStone
End Sub

Private Sub cmdClose_Click()
    DoCmd.Close acForm, Me.Name, acSavePrompt
    DoCmd.TransferText acExportDelim, , "tbl_PhotogrammetryRecords", CurrentProject.Path & "\Data_Database\PNA_2021_tbl_PhotogrammetryRecords.csv", True
End Sub


Private Sub cmdRemoveFilter_Click()
    ' removes any set filters
      On Error GoTo NoFilter
    DoCmd.RunCommand acCmdRemoveAllFilters
NoFilter:
    MsgBox "You have no filters set."
End Sub

'
'  RECORDS NUMBER SECTION
'
Private Sub Form_Current()
   Me.txtRecs = "Record" & " " & Form.CurrentRecord & " " & "of" & " " & Form.RecordsetClone.RecordCount
End Sub
'
'  UPDATE TIME AND DATE SECTION
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

