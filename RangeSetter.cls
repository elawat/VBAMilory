VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RangeSetter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public rngInstance As Range
Private dictNames As Scripting.Dictionary
 
Private Sub SetUpNamesDictionary()
Dim nm As Name

If Not dictNames Is Nothing Then Set dictNames = Nothing
Set dictNames = New Scripting.Dictionary
dictNames.CompareMode = TextCompare

For Each nm In ThisWorkbook.Names
    If Not (Strings.Right(nm.RefersTo, 5) = "#REF!") Then
        dictNames(nm.Name) = nm.RefersTo
    End If
Next

End Sub



Public Sub LoadFromAddress(ws As Worksheet, address As String)

On Error Resume Next
Set rngInstance = ws.Range(address)
On Error GoTo 0

End Sub
