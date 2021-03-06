VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UserParamGetter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public param As UserParam
Public errorMessage As String

Private Sub Class_Initialize()

Call Load
errorMessage = Validate

End Sub


Public Sub Load()

Set param = New UserParam

param.toolMode = GetToolMode()
param.miloryPath = GetMiloryPath()
param.DataPath = GetDataPath()
param.ResultPath = GetResultPath
param.CompareResuts = CompareDataInResAndData
param.SendToResults = SendToResults
param.SaveMiloryForEachObj = SaveMiloryPerObj
param.SaveAsPdf = SaveAsPdf
param.CreateSummary = CreateSummary
param.OutputFolderPath = GetOutputFolderPath
param.SelectObjectMode = GetSelectObjectMode

End Sub

Private Function Validate() As String

Dim msg As String

'--- Must be selected
If param.toolMode = 0 Then
    msg = "Wybierz tryb narz�dzia." & vbNewLine
End If

'--- Must be selected
If param.SelectObjectMode = 0 Then
    msg = msg & "Wybierz spos�b przeprowadzenia oblicze�." & vbNewLine
End If

'--- Cannot be empty and file needs to exist
If doesFileExist(param.miloryPath) = False Then
    msg = msg & "�cie�ka do arkusza Milory nie zosta�a podana lub jest niepoprawna." & vbNewLine
End If

If doesFileExist(param.DataPath) = False Then
    msg = msg & "�cie�ka do pliku 'Dane' nie zosta�a podana lub jest niepoprawna." & vbNewLine
End If

If param.toolMode = 1 And doesFileExist(param.ResultPath) = False Then
    msg = msg & "�cie�ka do pliku 'Wyniki' nie zosta�a podana lub jest niepoprawna." & vbNewLine
End If

If doesFolderExist(param.OutputFolderPath) = False Then
    msg = msg & "�cie�ka do folderu wyj�ciowego nie zosta�a podana lub jest niepoprawna." & vbNewLine
End If

If param.SendToResults = 0 And param.SaveMiloryForEachObj = 0 And param.SaveAsPdf = 0 _
    And param.CreateSummary = 0 Then
    msg = msg & "Wybierz rodzaj plik�w wyjciowych."
End If

Validate = msg

End Function


Private Function GetToolMode() As Integer

Dim dd As DropDown
Dim mode As Integer

Set dd = shUstawienia.DropDowns("DropDownMode")
mode = dd.value

GetToolMode = mode

End Function

Private Function GetSelectObjectMode() As Integer

Dim dd As DropDown

Set dd = shUstawienia.DropDowns("DropDownSelObj")

GetSelectObjectMode = dd.value

End Function

Private Function GetMiloryPath() As String

GetMiloryPath = shUstawienia.Range("MiloryPath").value

End Function

Private Function GetDataPath() As String

GetDataPath = shUstawienia.Range("InputDataPath").value

End Function

Private Function GetResultPath() As String

GetResultPath = shUstawienia.Range("GivenResultsFilePath").value

End Function

Private Function GetOutputFolderPath() As String

GetOutputFolderPath = shUstawienia.Range("OutputPath").value
If GetOutputFolderPath <> vbNullString And Right(GetOutputFolderPath, 1) <> "\" Then GetOutputFolderPath = GetOutputFolderPath & "\"


End Function

Private Function CompareDataInResAndData() As Boolean

CompareDataInResAndData = True

If shUstawienia.OLEObjects("cbxPorownaj").Object.value = 0 Then
    CompareDataInResAndData = False
End If

End Function


Private Function SendToResults() As Boolean

SendToResults = True

If shUstawienia.OLEObjects("cbxWyslijWyniki").Object.value = 0 Then
    SendToResults = False
End If

End Function

Private Function SaveMiloryPerObj() As Boolean

SaveMiloryPerObj = True

If shUstawienia.OLEObjects("cbxZapiszMilory").Object.value = 0 Then
    SaveMiloryPerObj = False
End If

End Function

Private Function SaveAsPdf() As Boolean

SaveAsPdf = True

If shUstawienia.OLEObjects("cbxPdf").Object.value = 0 Then
    SaveAsPdf = False
End If

End Function

Private Function CreateSummary() As Boolean

CreateSummary = True

If shUstawienia.OLEObjects("cbxMiloryzestawienie").Object.value = 0 Then
    CreateSummary = False
End If

End Function


Private Function CheckIfGenereteNewResultsFile() As Boolean

Dim dd As DropDown
Dim mode As Integer

Set dd = shUstawienia.DropDowns("DropDownMode")

mode = dd.value

If mode = 2 Then CheckIfGenereteNewResultsFile = True


End Function

Private Function CheckIfModeSlected() As Boolean

Dim dd As DropDown
Dim mode As Integer

CheckIfModeSlected = True


Set dd = shUstawienia.DropDowns("DropDownMode")

mode = dd.value

If mode = 0 Then CheckIfModeSlected = False


End Function



