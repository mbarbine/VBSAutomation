Dim fso, fldr, fc, f1 'fldname,  srcFile
set FSO = Wscript.CreateObject("scripting.FileSystemObject")
fldname = "" 
Const DateofFile=-730 

DeleteFiles = FSO.GetFolder(fldname)
Set fldr = fso.GetFolder(fldname)
 
Recurse fldr

on error resume next 
 
Set fldr = Nothing
Set fso = Nothing
Wscript.Quit
 
Public Sub Recurse( ByRef fldr)
dim subfolders,files,folder,file
Dim srcFile
Set subfolders = fldr.SubFolders
Set files = fldr.Files
 
For Each srcfile in files
If DateDiff("d", Now, srcFile.DateLastModified) < DateofFile Then
FSO.DeleteFile srcFile, True

End If
Next 

End sub
 


