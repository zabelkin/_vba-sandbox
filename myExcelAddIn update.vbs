On Error Resume Next

Dim objFSO
Set objFSO  = WScript.CreateObject("Scripting.FileSystemObject")

Dim objFile
Dim an, an_1, cand_path

fileFrom = "C:\Users\zabelkinav\Documents\проекты\_VBA sandbox\DEV myExcelAddIn.xlam"
fileTo = "C:\Users\zabelkinav\AppData\Roaming\Microsoft\Excel\XLSTART\myExcelAddIn.xlam"

Set objFile = objFSO.GetFile(fileTo)
objFile.Attributes = 0

objFSO.CopyFile fileFrom, fileTo, True

Set objFile = objFSO.GetFile(fileTo)
objFile.Attributes = 1

Set objFile = Nothing
Set objFSO  = Nothing

if not Err then
	Wscript.Echo "All Ok"
else 
	Wscript.Echo "Error:" & Chr(13) &Chr(13) & Err.Description
end if
