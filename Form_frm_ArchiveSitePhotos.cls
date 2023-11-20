VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frm_ArchiveSitePhotos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub txt_GoTo_AfterUpdate()
    On Error GoTo 0
    Me.Form.Filter = "File_Name=" & Me.txtSearch
    Me.Form.FilterOn = True
End Sub

Private Sub txt_Filter_AfterUpdate()
    On Error GoTo 0
    Me.Form.Filter = "[Record_Type]='" & Me.TypeFilter.Value & "'"
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
Private Sub cmdFilter_Click()
    Dim selectedSeason As Variant
    Dim selectedType As Variant
    Dim searchValue As String
    Dim columnName As String
    Dim filterCriteria As String
    
    selectedSeason = Me.SeasonFilter.Value
    selectedType = Me.TypeFilter.Value
    searchValue = Nz(Me.txtSearch.Value, "") ' Handle null value
    
    columnName = GetColumnName(Nz(Me.cboPrefix.Value, "")) ' Handle null value
    
    filterCriteria = ""
    
    If selectedSeason <> "" Then
        filterCriteria = "ID_Season = " & selectedSeason
    End If
    
    If selectedType <> "" Then
        If filterCriteria <> "" Then
            filterCriteria = filterCriteria & " AND "
        End If
        filterCriteria = filterCriteria & "Class_RecordClass = '" & selectedType & "'"
    End If
    
    If columnName <> "" And searchValue <> "" Then
        If filterCriteria <> "" Then
            filterCriteria = filterCriteria & " AND "
        End If
        filterCriteria = filterCriteria & columnName & " LIKE '*" & searchValue & "*'"
    End If
    
    If filterCriteria <> "" Then
        Me.Filter = filterCriteria
        Me.FilterOn = True
    Else
        Me.Filter = ""
        Me.FilterOn = False
    End If
End Sub
Private Function GetColumnName(prefix As String) As String
    Select Case prefix
        Case "TR": GetColumnName = "Link_Trench"
        Case "CTXT": GetColumnName = "Link_Context"
        Case "SMPL": GetColumnName = "Link_Sample"
        Case "SF": GetColumnName = "Link_SpecialFind"
        Case "BF": GetColumnName = "Link_Find"
        Case "M": GetColumnName = "Link_Masonry"
        Case Else: GetColumnName = ""
    End Select
End Function
Private Sub Form_AfterUpdate()
    On Error GoTo 0
End Sub
Private Sub cmdOpen_Click()
    Dim filePath As String
    filePath = Me.Photo_Path.Value ' Get the value from the Photo_Path text box
    
    If filePath <> "" Then
        On Error GoTo 0
        Application.FollowHyperlink filePath
    Else
        MsgBox "Photo path is empty.", vbExclamation, "No Path"
    End If
End Sub
Private Sub Form_Current()
    Me.txtRecs = "Record" & " " & Form.CurrentRecord & " " & "of" & " " & Form.RecordsetClone.RecordCount
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

