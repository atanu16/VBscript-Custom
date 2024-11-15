' VBScript to delete temp files and recreate temp directories

' Create FileSystemObject
Set fso = CreateObject("Scripting.FileSystemObject")

' Get the path of the user's temp directory and system temp directory
userTempDir = CreateObject("WScript.Shell").ExpandEnvironmentStrings("%%temp%%")
windowsTempDir = "C:\Windows\Temp"

' Function to delete all files in a directory
Sub DeleteFilesInDir(dirPath)
    If fso.FolderExists(dirPath) Then
        Set folder = fso.GetFolder(dirPath)
        For Each file In folder.Files
            On Error Resume Next ' In case a file is locked or cannot be deleted
            file.Delete True
            On Error GoTo 0
        Next
        ' Delete all subfolders as well
        For Each subFolder In folder.SubFolders
            On Error Resume Next ' In case a folder cannot be deleted
            subFolder.Delete True
            On Error GoTo 0
        Next
    End If
End Sub

' Function to recreate a directory
Sub RecreateDirectory(dirPath)
    If fso.FolderExists(dirPath) Then
        fso.DeleteFolder dirPath, True
    End If
    fso.CreateFolder dirPath
End Sub

' Main Script Execution
WScript.Echo "Cleaning up temporary files..."

' Delete files in the user's temp directory
WScript.Echo "Deleting files in " & userTempDir
DeleteFilesInDir userTempDir

' Delete files in the Windows Temp directory
WScript.Echo "Deleting files in " & windowsTempDir
DeleteFilesInDir windowsTempDir

' Recreate the temp directories
WScript.Echo "Recreating temp directories..."
RecreateDirectory userTempDir
RecreateDirectory windowsTempDir

WScript.Echo "Temporary files have been deleted and directories recreated."
