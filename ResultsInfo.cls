VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ResultsInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub Printer(errorMsg As String)
Dim rngComment As Range

Set rngComment = shUstawienia.Range("Komentarz")

If errorMsg = vbNullString Then
    rngComment.value = Format(Now, "dd/mm/yyyy hh:mm:ss")
    MsgBox "Pliki wyjœciowe zosta³y zapisane.", , "Ukoñczono obliczenia"
Else
    rngComment.value = Format(Now, "dd/mm/yyyy hh:mm:ss") & " " & errorMsg
    MsgBox errorMsg, , "SprawdŸ ustawienia"
End If

End Sub
 
