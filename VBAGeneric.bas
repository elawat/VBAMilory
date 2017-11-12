Attribute VB_Name = "VBAGeneric"
Option Explicit


Public Function doesFileExist(path As String) As Boolean

Dim fs As Scripting.FileSystemObject

Set fs = New Scripting.FileSystemObject
doesFileExist = fs.FileExists(path)

Set fs = Nothing

End Function


Public Function doesFolderExist(path As String) As Boolean

Dim fs As Scripting.FileSystemObject

Set fs = New Scripting.FileSystemObject
doesFolderExist = fs.FolderExists(path)

Set fs = Nothing

End Function

