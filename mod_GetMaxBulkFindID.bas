Public Function GetMaxBulkFindID() As Long
    Dim db As DAO.Database
    Dim rst As DAO.Recordset

    Set db = CurrentDb()
    Set rst = db.OpenRecordset("SELECT MAX(ID_BagID) AS MaxValue FROM tbl_FindsBulkRecords", dbOpenSnapshot)

    If Not (rst.EOF And rst.BOF) Then
        GetMaxBulkFindID = Nz(rst!MaxValue, 0)
    Else
        GetMaxBulkFindID = 0
    End If

    rst.Close
    Set rst = Nothing
    Set db = Nothing
End Function
