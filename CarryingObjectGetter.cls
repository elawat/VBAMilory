VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CarryingObjectGetter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public colObjects As Collection
Private map As Mapper
Private params As ToolParamGetter

Private Sub Class_Terminate()

If Not map Is Nothing Then Set map = Nothing
If Not params Is Nothing Then Set params = Nothing
      
End Sub

Public Sub Load(wbInput As Workbook)
Dim rowInMap As Object
Dim rowInInput As Integer
Dim lastrow As Long
Dim wsInput As Worksheet
Dim obj As CarryingObject
Dim value As Variant

Set map = New Mapper

Call map.Load("ObjFromDataFileMap")
Set params = New ToolParamGetter
Set wsInput = wbInput.Worksheets(params.toolParams.InputDataSheetIndex)
Set colObjects = New Collection

lastrow = wsInput.Cells(wsInput.Rows.Count, params.toolParams.InputFirstCol).End(xlUp).row

For rowInInput = params.toolParams.InputFirstDataRow To lastrow '--- Loop through all data rows in input files
    Set obj = New CarryingObject
    For Each rowInMap In map.MapCollection
        value = wsInput.Cells(rowInInput, rowInMap.columnIndex)
        Call obj.LoadProperty(obj, rowInMap.propertyName, value)
    Next rowInMap
    Call ValidateObj(obj)
    colObjects.add obj
    Set obj = Nothing
Next rowInInput

Set wsInput = Nothing

End Sub

Private Sub ValidateObj(obj As CarryingObject)

obj.IsValid = False

Call SetMainType(obj)
If obj.mainType <> vbNullString Then
    obj.IsValid = True
End If


End Sub

Private Sub SetMainType(obj As CarryingObject)
Dim typ As String
typ = vbNullString

If obj.ConstructionType = "B" Then
    If obj.BeamNo <> "" And IsNumeric(obj.BeamNo) Then
        If CInt(obj.BeamNo) = 2 Then
            typ = "d"
        End If
    Else
        If obj.kerb = 1 Then
            typ = "bk"
        Else
            typ = "b"
        End If
    End If
ElseIf obj.ConstructionType = "P" Then
    If obj.kerb = 1 Then
        typ = "pk"
    Else
        typ = "p"
    End If
End If

obj.mainType = typ


End Sub



