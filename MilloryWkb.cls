VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MilloryWkb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public wb As Workbook
Private staticWS As StaticSheet
Private maintypeWS As MainTypeSheet


Public Sub Load(wbMilloryPath As String)

Set wb = Workbooks.Open(wbMilloryPath, False, True)

End Sub

Private Sub Class_Terminate()

If Not wb Is Nothing Then
    wb.Close (False)
    Set wb = Nothing
End If
      
End Sub


Public Sub Fill(obj As CarryingObject)

Dim map As Mapper
Dim objMap As CObjMap
Dim rng As RangeSetter
Dim address As Variant


'--- Fill maintype worksheet
Set maintypeWS = New MainTypeSheet
Call maintypeWS.Load(obj.mainType, wb)
Set map = New Mapper
Call map.Load("ObjToMiloryMap", obj.mainType)
For Each objMap In map.MapCollection
    If Not objMap.MiloryAddress Is Nothing Then
        For Each address In objMap.MiloryAddress
            Set rng = New RangeSetter
            Call rng.LoadFromAddress(maintypeWS.ws, CStr(address))
            If Not rng.rngInstance Is Nothing Then
                rng.rngInstance.value = obj.GetPropValueFromPropName(obj, objMap.propertyName)
                Debug.Print rng.rngInstance.address, rng.rngInstance.value
            End If
            Set rng = Nothing
        Next address
    End If
Next objMap
Set map = Nothing

'--- Fill static sheet
Set staticWS = New StaticSheet
Call staticWS.Load(wb)
Set map = New Mapper
Call map.Load("ObjToMiloryMap", "all")
For Each objMap In map.MapCollection
    If Not objMap.MiloryAddress Is Nothing Then
        For Each address In objMap.MiloryAddress
            Set rng = New RangeSetter
            Call rng.LoadFromAddress(staticWS.ws, CStr(address))
            If Not rng.rngInstance Is Nothing Then
                rng.rngInstance.value = obj.GetPropValueFromPropName(obj, objMap.propertyName)
            End If
            Set rng = Nothing
        Next address
    End If
Next objMap
Set map = Nothing

'--- Fill results in obj
Call FillObjResultsFromMillory(obj)

obj.calcsDone = True

End Sub

Private Sub FillObjResultsFromMillory(obj As CarryingObject)

Dim map As Mapper
Dim objMap As CObjMap
Dim value As Variant
Dim rng As RangeSetter

Set map = New Mapper
Call map.Load("MilloryResultsToObjMap", obj.mainType)

For Each objMap In map.MapCollection
    If Not objMap.MiloryAddress Is Nothing Then
        Set rng = New RangeSetter
        Call rng.LoadFromAddress(maintypeWS.ws, CStr(objMap.MiloryAddress(1)))
        If Not rng.rngInstance Is Nothing Then
            If Not IsError(rng.rngInstance.value) Then
                If rng.rngInstance.value <> vbNullString Then
                    value = rng.rngInstance.value
                End If
            End If
        End If
        Set rng = Nothing
        Call obj.LoadProperty(obj, objMap.propertyName, value)
    End If
Next objMap

Set map = Nothing

End Sub

Public Sub Save(outputFolder As String, obj As CarryingObject)
Dim output As OutputFile
Dim path As String

Set output = New OutputFile
path = output.GetPath(outputFolder, obj)

wb.SaveAs filename:=path, FileFormat:=51

End Sub




Public Sub SaveAdPdf(obj As CarryingObject, wordSetter As WordDocSetter, _
            Optional outputFolder As String, Optional wordSummary As Document)
Dim output As OutputFile
Dim path As String
Dim miloryDetailsGetter As MilorySheetsDetailsGetter
Dim exporter As MiloryPdfExporter
Dim rng As RangeSetter
Dim colRanges As Collection
Dim milDetails As MilorySheetsDetails

Set exporter = New MiloryPdfExporter
Call exporter.Initialize(wordSetter)
Set miloryDetailsGetter = New MilorySheetsDetailsGetter
If outputFolder <> vbNullString Then
    Set output = New OutputFile
    path = output.GetPath(outputFolder, obj, "pdf")
End If
Set rng = New RangeSetter

Set milDetails = miloryDetailsGetter.dict("all")
Call rng.LoadFromAddress(staticWS.ws, milDetails.printRange)
If Not rng.rngInstance Is Nothing Then
    If colRanges Is Nothing Then
        Set colRanges = New Collection
    End If
    colRanges.add rng.rngInstance
    Set rng.rngInstance = Nothing
End If

Set milDetails = miloryDetailsGetter.dict(obj.mainType)
Call rng.LoadFromAddress(maintypeWS.ws, milDetails.printRange)
If Not rng.rngInstance Is Nothing Then
     If colRanges Is Nothing Then
        Set colRanges = New Collection
    End If
    colRanges.add rng.rngInstance
    Set rng.rngInstance = Nothing
End If

If Not colRanges Is Nothing Then
    Call exporter.PasteRangesAsPict(colRanges, wordSummary)
    If path <> vbNullString Then
        Call exporter.ExportToPdf(path)
    End If
    Set colRanges = Nothing
End If


Set exporter = Nothing
Set miloryDetailsGetter = Nothing
Set output = Nothing
Set rng = Nothing
Set milDetails = Nothing

End Sub
