' Module Name: mod_DatabaseInfoModule

Public Function GetDatabaseName() As String
    ' Define the database name
    GetDatabaseName = "ArcDatabase"
End Function

Public Function GetDatabaseVersion() As String
    ' Define the database version
    GetDatabaseVersion = "Version: Labëria (03.11.02b)"
End Function

Public Function GetKnownIssues() As String
    ' Define the known issues
    GetKnownIssues = "- Non Invasive: masonry records, building records, and photogrammetry records are not currently implimented in this version" & vbCrLf &
                     "- Invasive: human bone records are not currently implimented in this version" & vbCrLf &
                     "- Datanase: paper records are not currently implimented in this version" & vbCrLf &
                     "- Samples: when opned a dialogue appears asking for 'CurrentProject.Path' and an error with the Sample ID Field causing the naviation of records to break." & vbCrLf &
                     "- Datasheet view: Formatting of datasheet views needs sorting."
End Function

