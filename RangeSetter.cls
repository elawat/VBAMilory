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

Function Load(ws As Worksheet, address As String)

On Error Resume Next
Set rngInstance = ws.Range(address)
On Error GoTo 0

End Function
