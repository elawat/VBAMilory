Attribute VB_Name = "Main"
Option Explicit

Sub RunMiloryCalcs()
Dim process As MiloryCalcs

Application.ScreenUpdating = False

Call SetDevLayout

Set process = New MiloryCalcs
process.Run

Call SetInitialLayout

Set process = Nothing

Application.ScreenUpdating = True

End Sub


