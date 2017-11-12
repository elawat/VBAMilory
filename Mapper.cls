VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Mapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public MapCollection As Collection

Public Sub Load(mapType As String, Optional mainType As String)


Select Case mapType
    Case "ObjFromDataFileMap"
        Call GetObjFromDataFileMapper
    Case "ObjToMiloryMap"
        Call GetObjToMiloryMap(mainType)
    Case "MilloryResultsToObjMap"
        Call GetMilloryResultsToObjMap(mainType)
    Case "ObjToResultsMap"
        Call GetObjToResultsMap
    

End Select


End Sub

Private Sub GetObjFromDataFileMapper()
Dim co As CObjMap
Dim i As Integer
Dim mapArr() As Variant

Set MapCollection = New Collection

mapArr = shMapa.Range("ObjFromInputMap").value

For i = LBound(mapArr) To UBound(mapArr)
    Set co = New CObjMap
    co.columnIndex = mapArr(i, 1)
    co.propertyName = mapArr(i, 2)
    MapCollection.add co
    Set co = Nothing
Next i

End Sub


Private Sub GetObjToMiloryMap(mainType As String)
Dim co As CObjMap
Dim i As Integer
Dim mapAddressArr() As Variant
Dim mapCOArr() As Variant
Dim address As Variant

Set MapCollection = New Collection

mapAddressArr = shMapa.Range(mainType & "Address").value
mapCOArr = shMapa.Range("InputCOProp").value

For i = LBound(mapCOArr) To UBound(mapCOArr)
    Set co = New CObjMap
    For Each address In Split(mapAddressArr(i, 1), ",")
        If co.MiloryAddress Is Nothing Then Set co.MiloryAddress = New Collection
        co.MiloryAddress.add address
    Next address
    co.propertyName = mapCOArr(i, 1)
    MapCollection.add co
    Set co = Nothing
Next i

End Sub

Private Sub GetMilloryResultsToObjMap(mainType As String)
Dim co As CObjMap
Dim i As Integer
Dim mapAddressArr() As Variant
Dim mapCOArr() As Variant
Dim address As Variant

Set MapCollection = New Collection

mapAddressArr = shMapa.Range(mainType & "AddressMilory").value
mapCOArr = shMapa.Range("ResultCOProp").value

For i = LBound(mapCOArr) To UBound(mapCOArr)
    Set co = New CObjMap
    For Each address In Split(mapAddressArr(i, 1), ",")
        If co.MiloryAddress Is Nothing Then Set co.MiloryAddress = New Collection
        co.MiloryAddress.add address
    Next address
    co.propertyName = mapCOArr(i, 1)
    MapCollection.add co
    Set co = Nothing
Next i

End Sub



Private Sub GetObjToResultsMap()
Dim co As CObjMap
Dim i As Integer
Dim mapArr() As Variant

Set MapCollection = New Collection

mapArr = shMapa.Range("ObjToResultsMap").value

For i = LBound(mapArr) To UBound(mapArr)
    Set co = New CObjMap
    co.columnIndex = mapArr(i, 1)
    co.propertyName = mapArr(i, 2)
    MapCollection.add co
    Set co = Nothing
Next i

End Sub
