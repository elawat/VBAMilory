Attribute VB_Name = "UserInteface"
Option Explicit

Private Function GetPassword() As String
GetPassword = shParametry.Range("password")
End Function

Public Sub ChangeMode()

Dim ws As Worksheet
Dim dd As DropDown
Dim mode As Integer
Dim rngInPath As Range

Application.ScreenUpdating = False

Set ws = ThisWorkbook.Worksheets("Ustawienia")
Set dd = ws.DropDowns("DropDownMode")
Set rngInPath = ws.Range("GivenResultsFilePath")

mode = dd.value

If mode = 1 Then
    ws.Shapes("cbxPorownaj").Visible = True
    ws.Rows(rngInPath.row).EntireRow.Hidden = False
ElseIf mode = 2 Then
    ws.Shapes("cbxPorownaj").Visible = False
    ws.Rows(rngInPath.row).EntireRow.Hidden = True
End If

End Sub

Public Sub SetInitialLayout()
Dim ws As Worksheet
Dim password As String

password = GetPassword

'Application.DisplayFormulaBar = False

For Each ws In ThisWorkbook.Worksheets

    ws.Protect password:=password, UserInterFaceOnly:=True
    If ws.Name <> shUstawienia.Name Then
    
        ws.Visible = xlSheetVeryHidden
    
    End If
    
Next ws

shUstawienia.Range("MiloryPath").Locked = False
shUstawienia.Range("InputDataPath").Locked = False
shUstawienia.Range("GivenResultsFilePath").Locked = False
shUstawienia.Range("OutputPath").Locked = False

End Sub

Public Sub SetDevLayout()
Dim ws As Worksheet
Dim password As String

password = GetPassword

Application.DisplayFormulaBar = True

For Each ws In ThisWorkbook.Worksheets
      
      ws.Unprotect (password)
      ws.Visible = xlSheetVisible
    
Next ws

End Sub

Public Sub Unblock()
Application.ScreenUpdating = False
Dim checkPassword As Integer

checkPassword = ProvidedCorrectPassword
If checkPassword = 1 Then
    Call SetDevLayout
ElseIf checkPassword = 0 Then
    MsgBox "Niepoprawne has�o", vbOKOnly

End If
    
Application.ScreenUpdating = True
End Sub

Public Function ProvidedCorrectPassword() As Integer
Dim password As String
Dim userPassword As String
password = shParametry.Range("password")


userPassword = InputBox("Wprowad� has�o, aby odblokowa� narz�dzie", "Has�o")

If password = userPassword Then
    ProvidedCorrectPassword = 1
ElseIf userPassword = vbNullString Then
    ProvidedCorrectPassword = 2
Else
    ProvidedCorrectPassword = 0
End If



End Function
