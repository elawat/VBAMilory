VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WordDocSetter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public wdApp As Object
Private quitWordApp As Boolean

Private Sub Class_Terminate()

If Not wdApp Is Nothing Then
    
    Call ClearClipBoard
    If quitWordApp Then
        wdApp.Quit
    End If
    Set wdApp = Nothing
End If


End Sub

Public Sub Load(Optional makeInvisible As Boolean = False)

Dim wordApp As process
Set wordApp = New process

quitWordApp = False

If wordApp.IsRunning("WINWORD.exe") Then
    Set wdApp = GetObject(, "Word.Application")
Else
    Set wdApp = CreateObject("Word.Application")
    quitWordApp = True
End If

wdApp.DisplayAlerts = False

If makeInvisible Then
    wdApp.Visible = False
Else
    wdApp.Visible = True
End If

Set wordApp = Nothing


End Sub


Public Function GetNew() As Document
Dim objDoc As Document
Set objDoc = wdApp.Documents.add

Set GetNew = objDoc

Set objDoc = Nothing

End Function

Public Sub CloseDoc(objDoc As Document)

objDoc.Close (False)
Set objDoc = Nothing

End Sub


Private Sub ClearClipBoard()
Dim oData As New DataObject 'object to use the clipboard
oData.SetText Text:=Empty 'Clear
oData.PutInClipboard 'take in the clipboard to empty it
End Sub


