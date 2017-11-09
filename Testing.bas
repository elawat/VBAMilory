Attribute VB_Name = "Testing"
Option Explicit


Sub TestCarryingObjectGetter()

Dim outputPathFolder As String
Dim ws As Worksheet
Dim i As Integer
Dim wbData  As Workbook

i = 1

'--- Set up outputPath
outputPathFolder = "C:\Users\elawa\Desktop\Milory\Testing_" & Format(Now, "ddmmyyyy_hhmmss") & "\"
MkDir outputPathFolder

'--- Open and save copy of Data workbook
Set wbData = Workbooks.Open("C:\Users\elawa\Desktop\Milory\Testing\Dane.xlsx")
wbData.SaveAs "C:\Users\elawa\Desktop\Milory\Testing\Dane" & Format(Now, "ddmmyyyy_hhmmss")

'--- Get input data
Dim singleCObj As CarryingObject
Dim colCObjects As CarryingObjectGetter
Set colCObjects = New CarryingObjectGetter
Call colCObjects.Load(wbData)
wbData.Close (False)

Set ws = Workbooks.add.Worksheets(1)

'--- Loop through all object and execute required tasks
For Each singleCObj In colCObjects.colObjects
    If singleCObj.DoCalc Then
    ws.Range("A" & i) = singleCObj.lp
    ws.Range("B" & i) = singleCObj.JNI
    ws.Range("C" & i) = singleCObj.mainType
    ws.Range("D" & i) = singleCObj.BeamNo
    ws.Range("E" & i) = singleCObj.ConstructionType
    ws.Range("F" & i) = singleCObj.kerb
    ws.Range("G" & i) = singleCObj.IsValid
    i = i + 1
    End If
Next singleCObj
    
Set colCObjects = Nothing


End Sub


Sub TestMilloryWkb()

Dim outputPathFolder As String
Dim ws As Worksheet
Dim i As Integer
Dim wbData  As Workbook
Dim wbMilory As MilloryWkb

i = 1

'--- Set up outputPath
outputPathFolder = "C:\Users\elawa\Desktop\Milory\Testing_" & Format(Now, "ddmmyyyy_hhmmss") & "\"
MkDir outputPathFolder

'--- Open and save copy of Data workbook
Set wbData = Workbooks.Open("C:\Users\elawa\Desktop\Milory\Testing\Dane.xlsx")
wbData.SaveAs "C:\Users\elawa\Desktop\Milory\Testing\Dane" & Format(Now, "ddmmyyyy_hhmmss")

'--- Get input data
Dim singleCObj As CarryingObject
Dim colCObjects As CarryingObjectGetter
Set colCObjects = New CarryingObjectGetter
Call colCObjects.Load(wbData)
wbData.Close (False)

Set ws = Workbooks.add.Worksheets(1)

'--- Loop through all object and execute required tasks
For Each singleCObj In colCObjects.colObjects
    If singleCObj.DoCalc Then
        Set wbMilory = New MilloryWkb
        Call wbMilory.Load("C:\Users\elawa\Desktop\Milory\Testing\Milory.xls")
        Call wbMilory.Fill(singleCObj)
    End If
Next singleCObj
    
Set colCObjects = Nothing
Set wbMilory = Nothing

End Sub
