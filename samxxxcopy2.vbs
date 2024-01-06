Set fso = CreateObject("Scripting.FileSystemObject")
On Error Resume Next

' Retrieve the source and destination folders from command-line arguments
sourceFolder = WScript.Arguments.Item(0)
destinationFolder = WScript.Arguments.Item(1)

' Validate the provided folder paths
If Not fso.FolderExists(sourceFolder) Then
  WScript.Echo "Source folder not found: " & sourceFolder
  WScript.Quit
End If
If Not fso.FolderExists(destinationFolder) Then
  WScript.Echo "Destination folder not found: " & destinationFolder
  WScript.Quit
End If



' Specify the source and destination folders
'sourceFolder = "C:\SourceFolder" ' Replace with your actual source folder path
'destinationFolder = "C:\DestinationFolder" ' Replace with your actual destination folder path

' Iterate through each file in the source folder
For Each file In fso.GetFolder(sourceFolder).Files

  ' Check if the file already exists in the destination folder
  If fso.FileExists(destinationFolder & "\" & file.Name) Then

      ' Compare the modification dates of the source and destination files
      If file.DateLastModified > fso.GetFile(destinationFolder & "\" & file.Name).DateLastModified Then

          ' If the source file is newer, overwrite the destination file
          fso.CopyFile file.Path, destinationFolder & "\" & file.Name, True
          WScript.Echo file.Name & " was updated."

      End If

  End If

Next
If Err.Number <> 0 Then
    WScript.Echo "Error encountered: " & Err.Description
End If
WScript.Echo "File copying and updating completed."