Set fso = CreateObject("Scripting.FileSystemObject")


' Specify the source and destination folders
'sourceFolder = "F:\delete\1\Qemulin\qemu" ' Replace with your actual source folder path
'destinationFolder = "F:\delete\1\Qemulin\qemu1" ' Replace with your actual destination folder path

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

WScript.Echo "File copying and updating completed."