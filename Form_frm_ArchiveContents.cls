VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frm_ArchiveContents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub txt_GoTo_AfterUpdate()
    On Error GoTo 0
    Me.Form.Filter = "File_Name=" & Me.txt_GoTo
    Me.Form.FilterOn = True
End Sub

Private Sub txt_Filter_AfterUpdate()
    On Error GoTo 0
    Me.Form.Filter = "[Record_Type]='" & Me.txt_Filter.Value & "'"
    Me.Form.FilterOn = True
End Sub

Private Sub cmdRemoveFilter_Click()
    On Error GoTo NoFilter
    DoCmd.RunCommand acCmdRemoveAllFilters
    Me.Form.FilterOn = False

NoFilter:
    On Error GoTo 0
    MsgBox "You now have no filters set."
End Sub

Private Sub cmdAddNew_Click()
    On Error GoTo 0
    DoCmd.OpenForm "frm_ArchiveFolders", acNormal, "", "", acEdit, acNormal
End Sub

Private Sub cmdAddToList_Click()
    On Error GoTo 0
    Associated_Find_List = Associated_Find
    Associated_Trench_List = Associated_Trench
    Associated_Context_List = Associated_Context
    Associated_Skeleton_List = Associated_Skeleton
    Associated_Building_List = Associated_Building
    Associated_Masonry_List = Associated_Masonry
    Associated_WorkedStone_List = Associated_WorkedStone
    Associated_Photogrammetry_List = Associated_Photogrammetry
    Associated_Room_List = Associated_Room
    Associated_Survey_List = Associated_Survey
End Sub

Private Sub Form_AfterUpdate()
    On Error GoTo 0
End Sub

Private Sub cmdOpen_Click()
    On Error GoTo 0
    Application.FollowHyperlink ([CurrentProject].[Path] & [File_Path] & [File_Name] & [File_Extention])
End Sub

Private Sub viewerImage_DblClick(Cancel As Integer)
    On Error GoTo 0
    Application.FollowHyperlink ([CurrentProject].[Path] & [File_Path] & [File_Name] & [File_Extention])
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

