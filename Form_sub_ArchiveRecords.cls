VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_sub_ArchiveRecords"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub cmdOpen_Click()
    On Error GoTo cmdopen_Click_Err
        Application.FollowHyperlink Me.txtFilePath
cmdBackup_Click_Exit:
    Exit Sub
cmdopen_Click_Err:
    MsgBox "Issue with opening archival document, check information is correct and the document is in the local archive.", , "Issue with opening document"
        Resume cmdBackup_Click_Exit
End Sub

Private Sub Command429_Click()
    DoCmd.OpenForm "frm_ArchiveContents", acNormal, "", "", acEdit, acNormal
End Sub
