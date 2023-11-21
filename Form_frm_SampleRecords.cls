Private Sub Form_Load()
    ' Initialize visibility on form load
    ToggleSampleTypeButtons False
    SetTabVisibility False, False, False, False
    'Update Last Record Value
    Me.ID_LastRecord.Value = GetMaxSampleID()
End Sub
'
'  ------ SAMPLE VIEW SECTION ------------------------------------------------------------------------------------
'
Private Sub ToggleSampleTypeButtons(isSingleSample As Boolean)
    ' Toggle visibility of sample type buttons and related controls
    Me.SampleType_Wood.Visible = isSingleSample
    Me.SampleType_Charcoal.Visible = isSingleSample
    Me.SampleType_Misc.Visible = isSingleSample
    Me.Type_SampleType.Visible = isSingleSample
End Sub
Private Sub SetTabVisibility(woodVisible As Boolean, depositVisible As Boolean,
                             miscVisible As Boolean, charcoalVisible As Boolean)
    ' Toggle visibility of tabs
    Me.Tab_Details.Pages(0).Visible = woodVisible
    Me.Tab_Details.Pages(1).Visible = depositVisible
    Me.Tab_Details.Pages(2).Visible = charcoalVisible
    Me.Tab_Details.Pages(3).Visible = miscVisible
End Sub
'
'  SINGLE SAMPLE
'
Private Sub cmdSingleSample_Click()
    ' Show the SampleType buttons
    ToggleSampleTypeButtons True
End Sub
'
'  BULK SAMPLE
'
Private Sub cmdBulkSample_Click()
    ' Hide the SampleType buttons and show the "Details_Deposit" tab
    ToggleSampleTypeButtons False
    SetTabVisibility False, True, False, False
End Sub
'
'  WOOD TYPE SAMPLE
'
Private Sub SampleType_Wood_Click()
    ' Show "Details_Wood" tab
    SetTabVisibility True, False, False, False
End Sub
'
'  MISC TYPE SAMPLE
'
Private Sub SampleType_Misc_Click()
    ' Show "Details_Misc" tab
    SetTabVisibility False, False, False, True
End Sub
'
'  CHARCOAL TYPE SAMPLE
'
Private Sub SampleType_Charcoal_Click()
    ' Show "Details_Charcoal" tab
    SetTabVisibility False, False, True, False
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
        rs.FindFirst "[ID_SampleID] = " & recordID

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

