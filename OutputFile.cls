VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OutputFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public Function GetPath(folderPath As String, obj As CarryingObject, Optional ext As String = "xlsx")
Dim lp As String
Dim roadNo As String
Dim km As String
Dim town As String
Dim filename As String

lp = Replace(obj.lp, ".", "")
roadNo = obj.roadNo
km = obj.km
town = obj.TownName

If Right(folderPath, 1) <> "\" Then folderPath = folderPath & "\"

filename = lp & "_dw" & roadNo & "_" & km & "_" & town & "." & ext

GetPath = folderPath & filename

End Function
