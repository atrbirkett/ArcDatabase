'
'  ------ CUT/DEPOSIT RECORD VIEW ------------------------------------------------------------------------------------
'
Private Sub UpdatePageVisibility()
    Dim cutValues As Variant
    cutValues = Array("Cut", "Structural")

    Dim isVisible As Boolean
    isVisible = Not IsNull(Me.Context_Type.Value) And (Me.Context_Type.Value = cutValues(0) Or Me.Context_Type.Value = cutValues(1))

    ' Setting visibility based on the condition
    Me.Tab_Details.Pages(0).Visible = isVisible
    Me.Tab_Details.Pages(1).Visible = Not isVisible
End Sub
Private Sub Context_Type_AfterUpdate()
    UpdatePageVisibility()
End Sub
Private Sub Form_Load()
    ' Update the page visibility when the form loads
    UpdatePageVisibility()
End Sub
'
'  ------ RELATED RECORDS SECTION ------------------------------------------------------------------------------------
'
'   **Note a requery of the Listbox appeaars in cmd_GoToRecord_Click()**
Private Sub ListBox_RelatedPhotos_DblClick(Cancel As Integer)
    'opens the image using the path created in column 7
    Application.FollowHyperlink Me.ListBox_RelatedPhotos.Column(7)
End Sub
'
'   PHOTOGRAPHS
'
Private Sub cmd_GoPhotographs_Click()
    On Error GoTo ErrorHandler

    Dim selectedUID As Long

    ' Check if an item is selected in the listbox
    If Me.ListBox_RelatedPhotos.ListIndex <> -1 Then
        ' Assuming UID is a numeric (AutoNumber) field
        selectedUID = Me.ListBox_RelatedPhotos.Column(8)

        ' Open frm_ArchiveSitePhotos as a popup
        DoCmd.OpenForm "frm_ArchiveSitePhotos", WindowMode:=acDialog

        ' Find the record with the matching UID
        With Forms("frm_ArchiveSitePhotos").RecordsetClone
            .FindFirst "UID = " & selectedUID
            If Not .NoMatch Then
                Forms("frm_ArchiveSitePhotos").Bookmark = .Bookmark
            Else
                MsgBox "Record not found.", vbInformation
            End If
        End With
    Else
        MsgBox "Please select an item from the list.", vbInformation
    End If

    Exit Sub

ErrorHandler:
    MsgBox "An error occurred: " & Err.Description
End Sub
'
'   SPECIAL FINDS
'
Private Sub cmd_GoSpecialFinds_Click()
    On Error GoTo ErrorHandler

    Dim selectedID_FindID As Long

    ' Check if an item is selected in the listbox
    If Me.ListBox_RelatedSpecialFinds.ListIndex <> -1 Then
        ' Assuming ID_FindID is a numeric field
        selectedID_FindID = Me.ListBox_RelatedSpecialFinds.Column(8)

        ' Open frm_FindsSpecailRecords as a popup
        DoCmd.OpenForm "frm_FindsSpecailRecords", WindowMode:=acDialog

        ' Find the record with the matching ID_FindID
        With Forms("frm_FindsSpecailRecords").RecordsetClone
            .FindFirst "ID_FindID = " & selectedUID
            If Not .NoMatch Then
                Forms("frm_FindsSpecailRecords").Bookmark = .Bookmark
            Else
                MsgBox "Record not found.", vbInformation
            End If
        End With
    Else
        MsgBox "Please select an item from the list.", vbInformation
    End If

    Exit Sub

ErrorHandler:
    MsgBox "An error occurred: " & Err.Description
End Sub
'
'   SAMPLES
'
Private Sub cmd_GoSamples_Click()
    On Error GoTo ErrorHandler

    Dim selectedID_SampleID As Long

    ' Check if an item is selected in the listbox
    If Me.ListBox_RelatedSamples.ListIndex <> -1 Then
        ' Assuming ID_SampleID is a numeric field
        selectedID_SampleID = Me.ListBox_RelatedSamples.Column(8)

        ' Open tbl_SampleRecords as a popup
        DoCmd.OpenForm "tbl_SampleRecords", WindowMode:=acDialog

        ' Find the record with the matching ID_SampleID
        With Forms("tbl_SampleRecords").RecordsetClone
            .FindFirst "ID_SampleID = " & selectedUID
            If Not .NoMatch Then
                Forms("tbl_SampleRecords").Bookmark = .Bookmark
            Else
                MsgBox "Record not found.", vbInformation
            End If
        End With
    Else
        MsgBox "Please select an item from the list.", vbInformation
    End If

    Exit Sub

ErrorHandler:
    MsgBox "An error occurred: " & Err.Description
End Sub
'
'   PHOTOGRAMMETRY
'
Private Sub cmd_GoPhotogrammetry_Click()
    On Error GoTo ErrorHandler

    Dim selectedUID As Long

    ' Check if an item is selected in the listbox
    If Me.ListBox_RelatedPhotogrammetry.ListIndex <> -1 Then
        ' Assuming UID is a numeric (AutoNumber) field
        selectedUID = Me.ListBox_RelatedPhotogrammetry.Column(8)

        ' Open frm_PhotogrammetryRecords as a popup
        DoCmd.OpenForm "frm_PhotogrammetryRecords", WindowMode:=acDialog

        ' Find the record with the matching UID
        With Forms("frm_PhotogrammetryRecords").RecordsetClone
            .FindFirst "UID = " & selectedUID
            If Not .NoMatch Then
                Forms("frm_PhotogrammetryRecords").Bookmark = .Bookmark
            Else
                MsgBox "Record not found.", vbInformation
            End If
        End With
    Else
        MsgBox "Please select an item from the list.", vbInformation
    End If

    Exit Sub

ErrorHandler:
    MsgBox "An error occurred: " & Err.Description
End Sub
'
'  ------ FOOTERS SECTION ------------------------------------------------------------------------------------
'
'  FILTERS
'
Private Sub cmdRemoveFilter_Click()
    On Error Resume Next
    DoCmd.RunCommand acCmdRemoveAllFilters
    If Err.Number = 0 Then
        Me.FilterOn = False
        MsgBox "All filters have been removed."
    Else
        MsgBox "Error " & Err.Number & ": " & Err.Description, vbExclamation, "Filter Removal Error"
    End If
End Sub
Private Sub cmd_GoToRecord_Click()
    Dim recordID As Variant
    Dim rs As Recordset

    recordID = Me.txt_GoTo.Value

    If IsNumeric(recordID) Then
        Set rs = Me.RecordsetClone
        rs.FindFirst "[ID_Context] = " & recordID

        If rs.NoMatch Then
            MsgBox "Record not found.", vbInformation, "Search Result"
        Else
            Me.Bookmark = rs.Bookmark
        End If
    Else
        MsgBox "Please enter a valid numeric ID.", vbInformation, "Invalid Input"
    End If

    If Not rs Is Nothing Then
        Set rs = Nothing
    End If
End Sub
'
'  RECORD SELECTORS SECTION
'
Private Sub Form_Current()
    ' Requery the ListBox_RelatedPhotos
    Me.ListBox_RelatedPhotos.Requery
    Me.ListBox_RelatedSpecialFinds.Requery
    Me.ListBox_RelatedSamples.Requery
    Me.ListBox_RelatedPhotogrammetry.Requery

    'Displays Record Number
    On Error GoTo ErrorHandler

    ' Update record count display
    Me.txtRecs = "Record " & Me.CurrentRecord & " of " & Me.RecordsetClone.RecordCount

    Exit Sub

ErrorHandler:
    ' Handle potential errors, for example when recordset is not available
    Me.txtRecs = "Record information unavailable"
    ' Optionally, log error or display a message box
End Sub
Private Sub cmdFirst_Click()
    'Go to first record
    If Not (Me.Recordset.BOF) Then
        Me.Recordset.MoveFirst
    End If
End Sub
Private Sub cmdLast_Click()
    'Go to last record
    If Not (Me.Recordset.EOF) Then
        Me.Recordset.MoveLast
    End If
End Sub
Private Sub cmdNext_Click()
    'Go to next record
    If Not (Me.Recordset.EOF) Then
        Me.Recordset.MoveNext
        If Me.Recordset.EOF Then Me.Recordset.MoveLast
    End If
End Sub
Private Sub cmdPrevious_Click()
    'Go to previous record
    If Not (Me.Recordset.BOF) Then
        Me.Recordset.MovePrevious
        If Me.Recordset.BOF Then Me.Recordset.MoveFirst
    End If
End Sub
Private Sub cmdNew_Click()
    'Go to new record
    On Error GoTo ErrorHandler
    DoCmd.RunCommand acCmdRecordsGoToNew
    Exit Sub

ErrorHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Error"
End Sub
