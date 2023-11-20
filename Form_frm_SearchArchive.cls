VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frm_SearchArchive"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdOpen_Click()
    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "frm_ArchiveContents"
    
    stLinkCriteria = "[File_Name]=" & Me![SearchResults]
    DoCmd.OpenForm stDocName, , , stLinkCriteria, , 3
End Sub

Private Sub SearchFor_Change()
'Create a string (text) variable
    Dim vSearchString As String

'Populate the string variable with the text entered in the Text Box SearchFor
    vSearchString = SearchFor.Text

'Pass the value contained in the string variable to the hidden text box SrchText,
'that is used as the sear4ch criteria for the Query QRY_SearchAll
    SrchText.Value = vSearchString

'Requery the List Box to show the latest results for the text entered in Text Box SearchFor
    Me.SearchResults.Requery


'Tests for a trailing space and exits the sub routine at this point
'so as to preserve the trailing space, which would be lost if focus was shifted from Text Box SearchFor
    If Len(Me.SrchText) <> 0 And InStr(Len(SrchText), SrchText, " ", vbTextCompare) Then
        'Set the focus on the first item in the list box
            Me.SearchResults = Me.SearchResults.ItemData(1)
            Me.SearchResults.SetFocus
        'Requery the form to refresh the content of any unbound text box that might be feeding off the record source of  the List Box
            DoCmd.Requery
        'Returns the cursor to the the end of the text in Text Box SearchFor,
        'and restores trailing space lost when focus is shifted to the list box
            Me.SearchFor = vSearchString
            Me.SearchFor.SetFocus
            Me.SearchFor.SelStart = Me.SearchFor.SelLength
            
        Exit Sub
    End If

'Set the focus on the first item in the list box
    Me.SearchResults = Me.SearchResults.ItemData(1)
    Me.SearchResults.SetFocus

'Requery the form to refresh the content of any unbound text box that might be feeding off the record source of  the List Box
    DoCmd.Requery

'Returns the cursor to the the end of the text in Text Box SearchFor
    Me.SearchFor.SetFocus

    If Not IsNull(Len(Me.SearchFor)) Then
        Me.SearchFor.SelStart = Len(Me.SearchFor)
    End If
End Sub

Private Sub SearchResults_DblClick(Cancel As Integer)
    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "frm_ArchiveContents"
    
    stLinkCriteria = "[File_Name]=" & Me![SearchResults]
    DoCmd.OpenForm stDocName, , , stLinkCriteria, , 3
End Sub

