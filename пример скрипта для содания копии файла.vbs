On Error Resume Next

Dim objFSO
Set objFSO  = WScript.CreateObject("Scripting.FileSystemObject")

Dim objFile
Dim an, an_1, cand_path

an = "c:\\some_path\test.xlsb"
an_1 = "\\shared_drive\some_shared_path\test.xlsb"
cand_path = "c:\\path2\test2.xlsb"

Set objFile = objFSO.GetFile(an)
objFile.Attributes = 0 ' R/W
Set objFile = objFSO.GetFile(an_1)
objFile.Attributes = 0

' copy to -1
objFSO.CopyFile an, an_1, True

' copy from  candidate
objFSO.CopyFile cand_path, an, True

Set objFile = objFSO.GetFile(an)
objFile.Attributes = 1 ' R/O
Set objFile = objFSO.GetFile(an_1)
objFile.Attributes = 1


Set objFile = Nothing
Set objFSO  = Nothing


if not Err then
	Wscript.Echo "All Ok"
else 
	Wscript.Echo "Error:" & Chr(13) &Chr(13) & Err.Description
end if
