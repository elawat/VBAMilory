VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UserParam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public toolMode As Integer '2 generete new results; 1 fill results file

Public miloryPath As String
Public DataPath As String
Public ResultPath As String
Public OutputFolderPath As String

Public CompareResuts As Boolean

Public SelectObjectMode As Integer '1 all objects; 2 selected objects


Public SendToResults As Boolean
Public SaveMiloryForEachObj As Boolean
Public SaveAsPdf As Boolean
Public CreateSummary As Boolean
