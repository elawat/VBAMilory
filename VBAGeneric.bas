Attribute VB_Name = "VBAGeneric"
Option Explicit


Public Function doesFileExist(path As String) As Boolean

Dim fs As New Scripting.FileSystemObject
Dim functionResult As Boolean

doesFileExist = fs.FileExists(path)

Set fs = Nothing

End Function


Public Function doesFolderExist(path As String) As Boolean

Dim fs As New Scripting.FileSystemObject
Dim functionResult As Boolean

doesFolderExist = fs.FolderExists(path)

Set fs = Nothing

End Function

