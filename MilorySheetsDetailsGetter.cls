VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MilorySheetsDetailsGetter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public dict As Scripting.Dictionary

Private Sub Class_Initialize()

Call GetMiloryDetailsFromSheet

End Sub


Private Function GetMiloryDetailsFromSheet()
Dim details() As Variant
Dim obj As MilorySheetsDetails
Dim i As Integer

Set dict = New Scripting.Dictionary


details = shParametry.Range("MiloryDetails").value
For i = LBound(details, 1) To UBound(details, 1)
    Set obj = New MilorySheetsDetails
    obj.sheetName = details(i, 1)
    obj.printRange = details(i, 2)
    obj.sheetIndex = details(i, 3)
    obj.mainType = details(i, 4)
    obj.indexInMap = details(i, 5)
    dict.add obj.mainType, obj
    Set obj = Nothing
Next i


End Function
