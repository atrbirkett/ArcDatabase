Option Compare Database
Function GetCurrentDatabasePath() As String
    GetCurrentDatabasePath = CurrentDb.Name
End Function
Private Sub cmd_Adminon_Click()
    ' Enter admin mode
    EnterAdminMode
    DoCmd.Close acForm, "nav_splashscreen"
End Sub

Private Sub cmdBackup_Click()

    DoCmd.TransferText acExportDelim, , "tbl_BuildingRecords", CurrentProject.Path & "\Data_Database\PNA_2021_tbl_BuildingRecords.csv", True
    DoCmd.TransferText acExportDelim, , "tbl_BuildingRoomRecords", CurrentProject.Path & "\Data_Database\PNA_2021_tbl_BuildingRoomRecords.csv", True
    DoCmd.TransferText acExportDelim, , "tbl_BuildingRecords_Phases", CurrentProject.Path & "\Data_Database\PNA_2021_tbl_BuildingRecords_Phases.csv", True
    DoCmd.TransferText acExportDelim, , "tbl_EMRecords", CurrentProject.Path & "\Data_Database\PNA_2021_tbl_EMRecords.csv", True
    DoCmd.TransferText acExportDelim, , "tbl_FindsRecords", CurrentProject.Path & "\Data_Database\PNA_2021_tbl_FindsRecords.csv", True
    DoCmd.TransferText acExportDelim, , "tbl_Levels", CurrentProject.Path & "\Data_Database\PNA_2021_tbl_Levels.csv", True
    DoCmd.TransferText acExportDelim, , "tbl_MasonryRecords", CurrentProject.Path & "\Data_Database\PNA_2021_tbl_MasonryRecords.csv", True
    DoCmd.TransferText acExportDelim, , "tbl_PhotogrammetryRecords", CurrentProject.Path & "\Data_Database\PNA_2021_tbl_PhotogrammetryRecords.csv", True
    DoCmd.TransferText acExportDelim, , "tbl_SurveyNotes", CurrentProject.Path & "\Data_Database\PNA_2021_tbl_SurveyNotes.csv", True
    DoCmd.TransferText acExportDelim, , "tbl_WorkedStoneRecords", CurrentProject.Path & "\Data_Database\PNA_2021_tbl_WorkedStoneRecords.csv", True
    DoCmd.TransferText acExportDelim, , "tbl_TrenchRecords", CurrentProject.Path & "\Data_Database\PNA_2021_tbl_TrenchRecords.csv", True
    DoCmd.TransferText acExportDelim, , "tbl_ContextRecords", CurrentProject.Path & "\Data_Database\PNA_2021_tbl_ContextRecords.csv", True
    DoCmd.TransferText acExportDelim, , "tbl_SkeletonRecords", CurrentProject.Path & "\Data_Database\PNA_2021_tbl_SkeletonRecords.csv", True
    DoCmd.TransferText acExportDelim, , "tbl_SampleRecords", CurrentProject.Path & "\Data_Database\PNA_2021_tbl_SampleRecords.csv", True

    MsgBox "Backup to .csv compleated!"
End Sub

Private Sub cmd_Close_Click()
    ' Close the current form
    DoCmd.Close acForm, "nav_splashscreen"
End Sub

Private Sub Form_Load()
    ExitAdminMode
    Me.txtDatabasePath.Value = GetCurrentDatabasePath()
End Sub
