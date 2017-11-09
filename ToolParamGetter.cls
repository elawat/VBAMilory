VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ToolParamGetter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public toolParams As ToolParam

Private Sub Class_Initialize()
 
Set toolParams = New ToolParam
Call LoadToolParam

End Sub

Private Function LoadToolParam()

toolParams.pdfHeader = GetPdfHeader()
toolParams.SelObjColIndex = GetRunCalcsColumn()

toolParams.InputFirstCol = GetInputFirstCol()
toolParams.InputLastCol = GetInputLastCol()
toolParams.InputFirstDataRow = getInputFirstDataRow()
toolParams.InputDataSheetIndex = GetInputDataSheetIndex()
toolParams.ResultsSheetIndex = GetResultsSheetIndex()
toolParams.ResultsFirstDataCol = GetResultsFirstDataCol()
toolParams.ResultJNIColumn = GetResultJNIColumn()


End Function

Private Function GetPdfHeader() As String

GetPdfHeader = shParametry.Range("pdfHeader").value

End Function

Private Function GetRunCalcsColumn() As Integer

GetRunCalcsColumn = shParametry.Range("RunCalcsColumn").value

End Function

Private Function GetInputFirstCol() As Integer

GetInputFirstCol = shParametry.Range("FirstDataCol").value

End Function

Private Function GetInputLastCol() As Integer

GetInputLastCol = shParametry.Range("LastDataCol").value

End Function


Private Function getInputFirstDataRow() As Integer

getInputFirstDataRow = shParametry.Range("FirstDataRow").value

End Function

Function GetInputDataSheetIndex() As Integer

GetInputDataSheetIndex = shParametry.Range("DataSheetIndex").value

End Function

Function GetResultsFirstDataCol() As Integer

GetResultsFirstDataCol = shParametry.Range("FirstDatColResults").value

End Function


Function GetResultJNIColumn() As Integer

GetResultJNIColumn = shParametry.Range("ResultJNIColumn").value

End Function
'
'Function GetRoadNo() As Integer
'
'GetRoadNo = shMapowanie.Range("roadNoColumn").Value
'
'End Function
'
'Function GetKM() As Integer
'
'GetKM = shMapowanie.Range("kmColumn").Value
'
'End Function
'
'Function GetTown() As Integer
'
'GetTown = shMapowanie.Range("townColumn").Value
'
'End Function


'--- Parameters of output file

Function GetResultsSheetIndex() As Integer

GetResultsSheetIndex = shParametry.Range("ResultsSheetIndex").value

End Function

Function GetLastRowCheckColResults() As Integer

GetLastRowCheckColResults = shParametry.Range("LastRowCheckColResults").value

End Function

'Function GetJNIColIndexInResult() As Integer
'
'GetJNIColIndexInResult = shParametry.Range("ResultJNIColumn").value
'End Function




