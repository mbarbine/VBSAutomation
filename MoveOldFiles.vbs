Option Explicit 

On Error Resume Next

Dim fso, FileSet, Path, File, DDiff, Date1, Date2, DestPath

Path = ""
DestPath = "" 


FileSet = GetDirContents(Path) 

For each File in FileSet 
Set File = fso.GetFile(Path & "\" & File)
Date1 = File.DateLastModified 


Date2 = Now()

DDiff = Abs(DateDiff("h", Date1, Date2))

If DDiff >= 720 Then
If Not fso.FileExists(DestPath & File.Name) Then
File.Move DestPath
        'wscript.echo File.Name
on error resume next 
Else
wscript.echo "Unable to move file [" & File.Name & "].  A file by this name already exists in the target directory."
End If
End If

Next 

On Error resume next 

Function GetDirContents(FolderPath) 
Dim  FileCollection, aTmp(), i 
Set fso = CreateObject("Scripting.FileSystemObject") 
Set FileCollection = fso.GetFolder(FolderPath).Files 

Redim aTmp(FileCollection.count - 1) 
i = -1 

For Each File in FileCollection 
i = i + 1 
aTmp(i) = File.Name 
Next 

GetDirContents = aTmp 


End Function
