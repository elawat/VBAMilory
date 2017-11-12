VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CarryingObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'--- Validation
Public IsValid As Boolean
Public mainType As String
Public calcsDone As Boolean


'--- From input file
Public lp As Variant

Public roadNo As Variant
Public km As Variant
Public JNI As Variant
Public TownName As Variant

'Public Skr As Variant
Public objType As Variant 'Rodzaj obiektu w ci¹gu trasy
'Public ObjName As Variant
'Public YearBuilt As Integer

Public DesignNorm As Variant 'Norma projektowania
Public DesignClass As Variant 'Projektowana klasa obci¹¿enia
'Public TechState As Variant
'Public CurrentCarrying As Variant

Public ConstructionType As Variant 'P³yta (P) Belka (B)
Public BeamNo As Variant 'Liczba dŸwigarów
Public ArchNo As Variant
'Public ArchSpread As Variant
Public StaticScheme As Variant

Public Z1 As Variant
Public PB As Variant
Public Z2 As Variant
Public d1 As Variant
Public c1 As Variant
Public a1 As Variant
Public b As Variant
Public a2 As Variant
Public c2 As Variant
Public d2 As Variant
Public g As Variant
Public k1 As Variant
Public k2 As Variant

Public DoCalc As Variant

Public kerb As Variant 'krawêŸniki [TAK-1; NIE - puste]

Public L1 As Variant
Public L2 As Variant
Public L3 As Variant
Public L4 As Variant
Public L5 As Variant
Public L6 As Variant
Public L7 As Variant
Public L8 As Variant


Public LoadClassWheeled1 As Variant
Public LoadClassWheeled2 As Variant
Public LoadClassTracklaying1 As Variant
Public LoadClassTracklaying2 As Variant

Public AxisLoad As Variant
Public Comments As Variant
Public SignType As Variant
Public SignsAmount As Variant


Public Function GetPropValueFromPropName(obj As CarryingObject, propertyName As String) As Variant
Dim value As Variant

Select Case propertyName
    Case "lp"
        value = obj.lp
    Case "roadNo"
        value = obj.roadNo
    Case "km"
        value = obj.km
    Case "JNI"
        value = obj.JNI
    Case "TownName"
        value = obj.TownName
    Case "objType"
        value = obj.objType
    Case "DesignNorm"
        value = obj.DesignNorm
    Case "DesignClass"
        value = obj.DesignClass
    Case "ConstructionType"
        value = obj.ConstructionType
    Case "BeamNo"
         value = obj.BeamNo
    Case "ArchNo"
         value = obj.ArchNo
    Case "StaticScheme"
        value = obj.StaticScheme
    Case "Z1"
        value = obj.Z1
    Case "PB"
        value = obj.PB
    Case "Z2"
        value = obj.Z2
    Case "d1"
        value = obj.d1
    Case "c1"
        value = obj.c1
    Case "a1"
        value = obj.a1
    Case "b"
        value = obj.b
    Case "a2"
        value = obj.a2
    Case "c2"
        value = obj.c2
    Case "d2"
        value = obj.d2
    Case "g"
        value = obj.g
    Case "k1"
        value = obj.k1
    Case "k2"
        value = obj.k2
    Case "DoCalc"
        value = obj.DoCalc
    Case "kerb"
        value = obj.kerb
    Case "L1"
        value = obj.L1
    Case "L2"
        value = obj.L2
    Case "L3"
        value = obj.L3
    Case "L4"
        value = obj.L4
    Case "L5"
        value = obj.L5
    Case "L6"
        value = obj.L6
    Case "L7"
        value = obj.L7
    Case "L8"
        value = obj.L8
    Case "LoadClassWheeled1"
        value = obj.LoadClassWheeled1
    Case "LoadClassWheeled2"
        value = obj.LoadClassWheeled2
    Case "LoadClassTracklaying1"
        value = obj.LoadClassTracklaying1
    Case "LoadClassTracklaying2"
        value = obj.LoadClassTracklaying2
    Case "AxisLoad"
        value = obj.AxisLoad
    Case "Comments"
        value = obj.Comments
    Case "SignType"
        value = obj.SignType
    Case "SignsAmount"
        value = obj.SignsAmount
        
        
End Select

GetPropValueFromPropName = value

End Function

Public Sub LoadProperty(obj As CarryingObject, propertyName As String, value As Variant)

Select Case propertyName
    Case "lp"
        obj.lp = value
    Case "roadNo"
        obj.roadNo = value
    Case "km"
        obj.km = value
    Case "JNI"
        obj.JNI = value
    Case "TownName"
        obj.TownName = value
    Case "objType"
        obj.objType = value
    Case "DesignNorm"
        obj.DesignNorm = value
    Case "DesignClass"
        obj.DesignClass = value
    Case "ConstructionType"
        obj.ConstructionType = value
    Case "BeamNo"
        obj.BeamNo = value
    Case "ArchNo"
        obj.ArchNo = value
    Case "StaticScheme"
        obj.StaticScheme = value
    Case "Z1"
        obj.Z1 = value
    Case "PB"
        obj.PB = value
    Case "Z2"
        obj.Z2 = value
    Case "d1"
        obj.d1 = value
    Case "c1"
        obj.c1 = value
    Case "a1"
        obj.a1 = value
    Case "b"
        obj.b = value
    Case "a2"
        obj.a2 = value
    Case "c2"
        obj.c2 = value
    Case "d2"
        obj.d2 = value
    Case "g"
        obj.g = value
    Case "k1"
        obj.k1 = value
    Case "k2"
        obj.k2 = value
    Case "DoCalc"
        obj.DoCalc = value
    Case "kerb"
        obj.kerb = value
    Case "L1"
        obj.L1 = value
    Case "L2"
        obj.L2 = value
    Case "L3"
        obj.L3 = value
    Case "L4"
        obj.L4 = value
    Case "L5"
        obj.L5 = value
    Case "L6"
        obj.L6 = value
    Case "L7"
        obj.L7 = value
    Case "L8"
        obj.L8 = value
    Case "LoadClassWheeled1"
        obj.LoadClassWheeled1 = value
    Case "LoadClassWheeled2"
        obj.LoadClassWheeled2 = value
    Case "LoadClassTracklaying1"
        obj.LoadClassTracklaying1 = value
    Case "LoadClassTracklaying2"
        obj.LoadClassTracklaying2 = value
    Case "AxisLoad"
        obj.AxisLoad = value
    Case "Comments"
        obj.Comments = value
    Case "SignType"
        obj.SignType = value
    Case "SignsAmount"
        obj.SignsAmount = value
        
End Select



End Sub



