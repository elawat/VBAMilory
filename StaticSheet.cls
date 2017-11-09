VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StaticSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public MainTypeSheetIndex As Integer
Public MainTypePrintRange As String
Public MainTypeIndexInMap As Integer
Public ws As Worksheet


Public Sub Load(wbMilory As Workbook)

Dim sheetDetails As MilorySheetsDetailsGetter
Dim obj As MilorySheetsDetails

Set sheetDetails = New MilorySheetsDetailsGetter
Set obj = sheetDetails.dict("all")

MainTypePrintRange = obj.printRange
MainTypeSheetIndex = obj.sheetIndex
MainTypeIndexInMap = obj.indexInMap

Set ws = wbMilory.Worksheets(MainTypeSheetIndex)


End Sub

