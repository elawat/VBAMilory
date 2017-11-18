VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ResultsWkb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private wsResult As Worksheet
Private param As ToolParamGetter
Private doesFillExisting As Boolean

Private Sub Class_Terminate()

Set wsResult = Nothing
Set param = Nothing

End Sub


Public Sub Load(userInput As UserParamGetter, resultsSavePath As String, objs As Collection, _
                invalidObjs As ObjectsWithIssues)
Dim wbResult As Workbook

Set param = New ToolParamGetter


If userInput.param.toolMode = 2 Then
    '--- Create Result file
    shWynikiTemplate.Copy
    Set wsResult = ActiveSheet
    Set wbResult = wsResult.Parent
    wsResult.Name = "Wyniki"
    doesFillExisting = False
Else
    '--- Use existing result file
    Set wbResult = Workbooks.Open(userInput.param.ResultPath)
    Set wsResult = wbResult.Worksheets(param.toolParams.ResultsSheetIndex)
    doesFillExisting = True
End If
wbResult.SaveAs resultsSavePath

Call Fill(objs, userInput.param.toolMode, _
            invalidObjs, userInput.param.CompareResuts)
wbResult.Close (True)

Set wbResult = Nothing

End Sub


Private Sub Fill(objs As Collection, toolMode As Integer, _
        issues As ObjectsWithIssues, Optional CompareWithExisting As Boolean = True)
Dim map As Mapper
Dim objMap As CObjMap
Dim value As Variant
Dim row As Integer
Dim oldvalue As Variant
Dim obj As CarryingObject

Set map = New Mapper
Call map.Load("ObjToResultsMap")
If toolMode = 2 Then CompareWithExisting = False

For Each obj In objs

    If obj.calcsDone Then
        
        row = 0
        row = GetRow(obj, toolMode)
        
        If row <> 0 Then
        
            For Each objMap In map.MapCollection
                    value = obj.GetPropValueFromPropName(obj, objMap.propertyName)
                    oldvalue = wsResult.Cells(row, objMap.columnIndex).Value2
                    wsResult.Cells(row, objMap.columnIndex).value = value
                    If CompareWithExisting And Trim(oldvalue) <> Trim(value) Then
                        wsResult.Cells(row, objMap.columnIndex).AddComment CStr(oldvalue)
                    End If
                    value = vbNullString
                    oldvalue = vbNullString
            Next objMap
        
        Else
        
            issues.AddNotFound obj
        
        End If
        
    End If


NextObj:

Next obj

Set map = Nothing

End Sub


Private Function GetRow(obj As CarryingObject, toolMode As Integer) As Integer
Dim row As Integer

If toolMode = 2 Then
    row = wsResult.Cells(wsResult.Rows.Count, param.toolParams.ResultsFirstDataCol).End(xlUp).row + 1
    If row < 5 Then row = 5
Else
    If Not IsError(Application.Match(Trim(obj.JNI), wsResult.Columns(param.toolParams.ResultJNIColumn), 0)) Then
        row = Application.Match(Trim(obj.JNI), wsResult.Columns(param.toolParams.ResultJNIColumn), 0)
    End If
End If

GetRow = row

End Function
