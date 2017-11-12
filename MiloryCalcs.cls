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


Public Sub Run()
Dim userInput As UserParamGetter
Dim timeStamp As Date
Dim outputPathFolder As String
Dim wbData  As Workbook
Dim wbMilory As MilloryWkb
'Dim doesFillExisting As Boolean
Dim wbResults As ResultsWkb
Dim resultSavePath As String
Dim wordSetter As WordDocSetter
Dim wordSummary As Document
Dim singleCObj As CarryingObject
Dim colCObjects As CarryingObjectGetter
Dim invalidObjs As ObjectsWithIssues
Dim resInfo As ResultsInfo

Application.ScreenUpdating = False
Application.DisplayAlerts = False

'--- Get initial user parameters
Set userInput = New UserParamGetter
'--- Validate user inputs
If userInput.errorMessage <> vbNullString Then GoTo ExitAndTidyUp

'--- Set up outputPath
timeStamp = Now
outputPathFolder = userInput.param.OutputFolderPath & "Wyniki_" & Format(timeStamp, "ddmmyyyy_hhmmss") & "\"
MkDir outputPathFolder

'--- Open and save copy of Data workbook
Set wbData = Workbooks.Open(userInput.param.DataPath)
wbData.SaveAs outputPathFolder & "Dane_" & Format(timeStamp, "ddmmyyyy_hhmmss")
'--- Get input data
Set invalidObjs = New ObjectsWithIssues
Set colCObjects = New CarryingObjectGetter
Call colCObjects.Load(wbData)
wbData.Close (False)
Set wbData = Nothing

'--- Set up word that it is used in pdf exporter
Set wordSetter = New WordDocSetter
Call wordSetter.Load

'--- Create summary pdf
If userInput.param.CreateSummary Then
    Set wordSummary = wordSetter.GetNew
End If

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

Application.StatusBar = "Saving outputs..."

'--- Save summery when needed
If Not wordSummary Is Nothing Then
    wordSummary.ExportAsFixedFormat OutputFileName:=outputPathFolder & "Milory_zestawienie.pdf", _
     ExportFormat:=wdExportFormatPDF
    wordSummary.Close (False)
    Set wordSummary = Nothing
End If

'--- Send to results
If userInput.param.SendToResults Then
    resultSavePath = outputPathFolder & "Wyniki_" & Format(timeStamp, "ddmmyyyy_hhmmss")
    Set wbResults = New ResultsWkb
    Call wbResults.Load(userInput, resultSavePath, colCObjects.colObjects, invalidObjs)
    Set wbResults = Nothing
End If

Call invalidObjs.PrintcolInvalidObjects
Call invalidObjs.PrintcolNotFoundInResultsObj

ExitAndTidyUp:
Set resInfo = New ResultsInfo
resInfo.Printer (userInput.errorMessage)

Set userInput = Nothing
Set invalidObjs = Nothing
Set colCObjects = Nothing
Set wordSetter = Nothing

Application.StatusBar = False
Application.ScreenUpdating = True
Application.DisplayAlerts = True

End Sub
