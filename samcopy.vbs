Option Explicit

Dim sourceFolder, destinationFolder
sourceFolder = "F:\delete\1\Qemulin\qemu1"
destinationFolder = "F:\delete\1\Qemulin\qemu"

CopyAndUpdateFiles sourceFolder, destinationFolder

Sub CopyAndUpdateFiles(sourcePath, destinationPath)
    Dim fso, sourceFiles, sourceFile, destinationFile

    ' Create FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' Check if source folder exists
    If fso.FolderExists(sourcePath) Then
        ' Check if destination folder exists, create if not
        If Not fso.FolderExists(destinationPath) Then
            fso.CreateFolder destinationPath
        End If

        ' Get the collection of files in the source folder
        Set sourceFiles = fso.GetFolder(sourcePath).Files

        ' Iterate through each file in the source folder
        For Each sourceFile In sourceFiles
            ' Construct the destination file path
            destinationFile = fso.BuildPath(destinationPath, fso.GetFileName(sourceFile))

            ' Check if the file already exists in the destination folder
            If fso.FileExists(destinationFile) Then
                ' If it exists, overwrite the existing file
                fso.CopyFile sourceFile, destinationFile, True
                WScript.Echo "Updated file: " & destinationFile
            End If
        Next

        WScript.Echo "File copy and update completed."
    Else
        WScript.Echo "Source folder does not exist."
    End If
End Sub
