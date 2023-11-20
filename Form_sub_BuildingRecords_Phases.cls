VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_sub_BuildingRecords_Phases"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub ID_Building_AfterUpdate()
    DoCmd.Requery "Associated_Room"
End Sub

Private Sub Dating_Period_AfterUpdate()
Dating_MinYear = [Dating_Period].[Column](1)
Dating_MaxYear = [Dating_Period].[Column](2)
End Sub

