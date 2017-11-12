VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ObjectsWithIssues"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public colInvalidObjects As Collection
Public colNotFoundInResultsObj As Collection


Public Sub AddInvalid(obj As CarryingObject)

If colInvalidObjects Is Nothing Then
    Set colInvalidObjects = New Collection
End If

colInvalidObjects.add obj


End Sub

Public Sub AddNotFound(obj As CarryingObject)

If colNotFoundInResultsObj Is Nothing Then
    Set colNotFoundInResultsObj = New Collection
End If

colNotFoundInResultsObj.add obj

End Sub

Public Sub PrintcolInvalidObjects()

Dim lastrow As Long
Dim rng As Range
Dim obj As CarryingObject
Dim i As Long

Set rng = shUstawienia.Range("invalidtypeObjects")

lastrow = shUstawienia.Cells(shUstawienia.Rows.Count, rng.Column).End(xlUp).row
If lastrow > rng.row Then
    shUstawienia.Range(shUstawienia.Cells(rng.row + 1, rng.Column), shUstawienia.Cells(lastrow, rng.Column)).ClearContents
End If

i = 1
If Not colInvalidObjects Is Nothing Then

    For Each obj In colInvalidObjects
        rng.Offset(i, 0).value = obj.JNI & ", " & obj.TownName
        i = i + 1
    Next obj
Else
    rng.Offset(i, 0).value = "Nie zarejestrowano"
End If

Set rng = Nothing

End Sub


Public Sub PrintcolNotFoundInResultsObj()

Dim lastrow As Long
Dim rng As Range
Dim obj As CarryingObject
Dim i As Long

Set rng = shUstawienia.Range("notfoundObjects")

lastrow = shUstawienia.Cells(shUstawienia.Rows.Count, rng.Column).End(xlUp).row
If lastrow > rng.row Then
    shUstawienia.Range(shUstawienia.Cells(rng.row + 1, rng.Column), shUstawienia.Cells(lastrow, rng.Column)).ClearContents
End If

i = 1
If Not colNotFoundInResultsObj Is Nothing Then
    
    For Each obj In colNotFoundInResultsObj
        rng.Offset(i, 0).value = obj.JNI & ", " & obj.TownName
        i = i + 1
    Next obj
Else
    rng.Offset(1, 0).value = "Nie zarejestrowano"
End If

Set rng = Nothing

End Sub



Private Sub Class_Terminate()

If Not colNotFoundInResultsObj Is Nothing Then Set colNotFoundInResultsObj = Nothing
If Not colInvalidObjects Is Nothing Then Set colInvalidObjects = Nothing

End Sub


