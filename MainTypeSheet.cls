VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MainTypeSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public MainTypeSheetIndex As Integer
Public MainTypeIndexInMap As Integer
Public MainTypePrintRange As String
Public ws As Worksheet


Public Sub Load(mainType As String, wbInput As Workbook)

Dim sheetDetails As MilorySheetsDetailsGetter
Dim obj As MilorySheetsDetails

Set sheetDetails = New MilorySheetsDetailsGetter
Set obj = sheetDetails.dict(mainType)

MainTypePrintRange = obj.printRange
MainTypeSheetIndex = obj.sheetIndex
MainTypeIndexInMap = obj.indexInMap

Set ws = wbInput.Worksheets(MainTypeSheetIndex)


End Sub



