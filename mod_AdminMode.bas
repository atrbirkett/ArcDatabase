' Module: mod_AdminMode
Public IsAdminMode As Boolean ' Declare a global variable to track admin mode

' Procedure to Exit admin mode
Public Sub ExitAdminMode()
    Call DoCmd.ShowToolbar("Ribbon", acToolbarNo)
    Call DoCmd.NavigateTo("acNavigationCategoryObjectType")
    Call DoCmd.RunCommand(acCmdWindowHide)

    ' Set the global variable to indicate admin mode
    IsAdminMode = True
End Sub

' Procedure to Enter admin mode
Public Sub EnterAdminMode()
    Call DoCmd.ShowToolbar("Ribbon", acToolbarYes)
    Call DoCmd.SelectObject(acTable, , True)

    ' Set the global variable to indicate admin mode
    IsAdminMode = False
End Sub