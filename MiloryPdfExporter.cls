VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MiloryPdfExporter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private objDoc As Document
Private params As ToolParamGetter
Private wordApp As Object
Private header As String

Public Sub Initialize(wordSetter As WordDocSetter)


Set params = New ToolParamGetter
Set objDoc = wordSetter.GetNew
Set wordApp = wordSetter.wdApp

End Sub

Private Sub Class_Terminate()

If Not objDoc Is Nothing Then
    objDoc.Close (False)
    Set objDoc = Nothing
End If
If Not params Is Nothing Then Set params = Nothing
If Not wordApp Is Nothing Then Set wordApp = Nothing

End Sub


Public Sub PasteRangesAsPict(ranges As Collection, Optional wordSummary As Document)

Dim rng As Range
Dim lastParag As Long

lastParag = objDoc.Paragraphs.Count
header = params.toolParams.pdfHeader

objDoc.Paragraphs(lastParag).Range.Text = header
'objDoc.Paragraphs(lastParag).Range.Style = objDoc.Styles("Heading 3")
objDoc.Paragraphs(lastParag).Range.Font.ColorIndex = wdBlack
objDoc.Paragraphs(lastParag).Alignment = wdAlignParagraphCenter

objDoc.Paragraphs.add
objDoc.Paragraphs.add
wordApp.Selection.MoveDown Unit:=wdLine, Count:=2

For Each rng In ranges
    objDoc.Paragraphs.add
    Call PasteRangeAsPict(rng)
Next rng

Call DeleteBlankPages(objDoc)


If Not wordSummary Is Nothing Then
    Call AddToSummary(wordSummary)
End If

End Sub

Private Function DeleteBlankPages(wd As Word.Document)

Dim par As Paragraph
For Each par In wd.Paragraphs
    If Len(par.Range.Text) <= 1 Then
        par.Range.Delete
    End If
Next par

End Function

Private Sub PasteRangeAsPict(rng)

rng.CopyPicture Appearance:=xlScreen, Format:=xlPicture
wordApp.Selection.Paste
wordApp.Selection.TypeParagraph


End Sub

Public Sub ExportToPdf(filePath)


objDoc.ExportAsFixedFormat OutputFileName:=filePath, _
     ExportFormat:=wdExportFormatPDF


End Sub

Private Sub AddToSummary(summaryWord As Document)

objDoc.Activate
objDoc.Parent.Selection.WholeStory
objDoc.Parent.Selection.Expand wdParagraph
objDoc.Parent.Selection.Copy

summaryWord.Activate
If InStr(summaryWord.Paragraphs(1).Range.Text, header) > 0 Then '--- To check if already pasted at least once
    summaryWord.Parent.Selection.EndKey
End If
summaryWord.Parent.Selection.PasteAndFormat wdPasteDefault
summaryWord.Parent.Selection.TypeBackspace

End Sub



