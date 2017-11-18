VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MiloryCalcs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private userInput As UserParamGetter
Private outputPathFolder As String
Private timeStamp As Date
Private invalidObjs As ObjectsWithIssues
Private colCObjects As CarryingObjectGetter
Private wordSetter As WordDocSetter
Private wordSummary As Document


Private Sub Class_Terminate()

If Not userInput Is Nothing Then Set userInput = Nothing
If Not invalidObjs Is Nothing Then Set invalidObjs = Nothing
If Not colCObjects Is Nothing Then Set colCObjects = Nothing
If Not wordSetter Is Nothing Then Set wordSetter = Nothing
If Not wordSummary Is Nothing Then Set wordSummary = Nothing

End Sub


Public Sub Run()

Application.ScreenUpdating = False
Application.DisplayAlerts = False

Call GetUserParameters
'--- Validate user inputs
If userInput.errorMessage = vbNullString Then

    Call SetUpOutputPath
    Call GetInputData
    Call SetUpWord
    Call CreateSummaryPdf
    Call WriteOutput
    Call SaveSummaryPdf
    Call PrintInvalid

End If

Call PrintResultsInfo

Application.StatusBar = False
Application.ScreenUpdating = True
Application.DisplayAlerts = True

End Sub

Private Sub GetUserParameters()

'--- Get initial user parameters
Set userInput = New UserParamGetter

End Sub

Private Sub SetUpOutputPath()

timeStamp = Now
outputPathFolder = userInput.param.OutputFolderPath & "Wyniki_" & Format(timeStamp, "ddmmyyyy_hhmmss") & "\"
MkDir outputPathFolder

End Sub

Private Sub GetInputData()
Dim wbData  As Workbook

'--- Open and save copy of Data workbook
Set wbData = Workbooks.Open(userInput.param.DataPath)
wbData.SaveAs outputPathFolder & "Dane_" & Format(timeStamp, "ddmmyyyy_hhmmss")

'--- Get input data
Set invalidObjs = New ObjectsWithIssues
Set colCObjects = New CarryingObjectGetter
Call colCObjects.Load(wbData)


wbData.Close (False)
Set wbData = Nothing

End Sub

Private Sub SetUpWord()

'--- Set up word that it is used in pdf exporter
Set wordSetter = New WordDocSetter
Call wordSetter.Load

End Sub

Private Sub CreateSummaryPdf()

'--- Create summary pdf
If userInput.param.CreateSummary Then
    Set wordSummary = wordSetter.GetNew
End If

End Sub

Private Sub WriteOutput()

Dim singleCObj As CarryingObject
Dim wbMilory As MilloryWkb

'--- Loop through all object and execute required tasks
For Each singleCObj In colCObjects.colObjects
    If singleCObj.DoCalc = 1 Or userInput.param.SelectObjectMode = 1 Then
    
        Application.StatusBar = "JNI " & singleCObj.JNI & " " & singleCObj.TownName
        
        If singleCObj.IsValid Then
            '--- Run Calcs and exports
            Set wbMilory = New MilloryWkb
            Call wbMilory.Load(userInput.param.miloryPath)
            Call wbMilory.Fill(singleCObj)
            
            '--- Save Milory
            If userInput.param.SaveMiloryForEachObj Then
                Call wbMilory.Save(outputPathFolder, singleCObj)
            End If
            
            '--- Save as pdf
            If userInput.param.SaveAsPdf Then
                Call wbMilory.SaveAdPdf(singleCObj, wordSetter, outputPathFolder, wordSummary)
            ElseIf userInput.param.CreateSummary Then
                Call wbMilory.SaveAdPdf(singleCObj, wordSetter, , wordSummary)
            End If
            
            Set wbMilory = Nothing
        Else
            Call invalidObjs.AddInvalid(singleCObj)
        End If
    
    End If
Next singleCObj

End Sub

Private Sub SaveSummaryPdf()

Application.StatusBar = "Saving outputs..."

'--- Save summery when needed
If Not wordSummary Is Nothing Then
    wordSummary.ExportAsFixedFormat OutputFileName:=outputPathFolder & "Milory_zestawienie.pdf", _
     ExportFormat:=wdExportFormatPDF
    wordSummary.Close (False)
    
End If


End Sub

Private Sub SendToResults()

Dim wbResults As ResultsWkb
Dim resultSavePath As String

'--- Send to results
If userInput.param.SendToResults Then
    resultSavePath = outputPathFolder & "Wyniki_" & Format(timeStamp, "ddmmyyyy_hhmmss")
    Set wbResults = New ResultsWkb
    Call wbResults.Load(userInput, resultSavePath, colCObjects.colObjects, invalidObjs)
    Set wbResults = Nothing
End If

End Sub

Private Sub PrintInvalid()

Call invalidObjs.PrintcolInvalidObjects
Call invalidObjs.PrintcolNotFoundInResultsObj

End Sub

Private Sub PrintResultsInfo()
Dim resInfo As ResultsInfo

Set resInfo = New ResultsInfo
resInfo.Printer (userInput.errorMessage)

End Sub
