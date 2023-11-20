Attribute VB_Name = "modRefreshTableLinks"
Option Compare Database


'*******************************************************************
'*  This module refreshes the links to any linked tables  *
'*******************************************************************


'Procedure to relink tables from the Common Access Database
Public Function RefreshTableLinks() As String

On Error GoTo ErrHandler
    Dim strEnvironment As String
    strEnvironment = GetEnvironment

    Dim db As DAO.Database
    Dim tdf As DAO.TableDef

    Dim strCon As String
    Dim strBackEnd As String
    Dim strMsg As String

    Dim intErrorCount As Integer

    Set db = CurrentDb

    'Loop through the TableDefs Collection.
    For Each tdf In db.TableDefs

            'Verify the table is a linked table.
            If Left$(tdf.Connect, 10) = ";DATABASE=" Then

                'Get the existing Connection String.
                strCon = Nz(tdf.Connect, "")

                'Get the name of the back-end database using String Functions.
                strBackEnd = Right$(strCon, (Len(strCon) - (InStrRev(strCon, "\") - 1)))

                'Debug.Print strBackEnd

                'Verify we have a value for the back-end
                If Len(strBackEnd & "") > 0 Then

                    'Set a reference to the TableDef Object.
                    Set tdf = db.TableDefs(tdf.Name)

                    If strBackEnd = "\Common Shares_Data.mdb" Or strBackEnd = "\Adverse Events.mdb" Then
                        'Build the new Connection Property Value - below needs to be changed to a constant
                        tdf.Connect = ";DATABASE=" & strEnvironment & strBackEnd
                    Else
                        tdf.Connect = ";DATABASE=" & CurrentProject.Path & strBackEnd

                    End If

                    'Refresh the table links
                    tdf.RefreshLink

                End If

            End If

    Next tdf

ErrHandler:

 If Err.Number <> 0 Then

    'Create a message box with the error number and description
    MsgBox ("Error Number: " & Err.Number & vbCrLf & _
            "Error Description: " & Err.Description & vbCrLf)

End If

End Function





