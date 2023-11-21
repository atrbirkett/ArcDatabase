Public Function GetMaxSampleID() As Long
    Dim db As DAO.Database
    Dim rst As DAO.Recordset

    Set db = CurrentDb()
    Set rst = db.OpenRecordset("SELECT MAX(ID_SampleID) AS MaxValue FROM tbl_SampleRecords", dbOpenSnapshot)

    If Not (rst.EOF And rst.BOF) Then
        GetMaxSampleID = Nz(rst!MaxValue, 0)
    Else
        GetMaxSampleID = 0
    End If

    rst.Close
    Set rst = Nothing
    Set db = Nothing
End Function
