Public Function GetMaxSpecialFindID() As Long
    Dim db As DAO.Database
    Dim rst As DAO.Recordset

    Set db = CurrentDb()
    Set rst = db.OpenRecordset("SELECT MAX(ID_BagID) AS MaxValue FROM tbl_FindsSpecialRecords", dbOpenSnapshot)

    If Not (rst.EOF And rst.BOF) Then
        GetMaxSpecialFindID = Nz(rst!MaxValue, 0)
    Else
        GetMaxSpecialFindID = 0
    End If

    rst.Close
    Set rst = Nothing
    Set db = Nothing
End Function
