On Error Resume Next

Dim objFSO
Set objFSO  = WScript.CreateObject("Scripting.FileSystemObject")

Dim objFile
Dim an, an_1, cand_path

an = "\\ncm.lo\FileServer\FileServer\ltdm\DB_Analog\Analog.xlsb"
an_1 = "\\ncm.lo\FileServer\FileServer\ltdm\DB_Analog\Analog-1.xlsb"
cand_path = "\\srv-08\exchange\Analyze\Конкурентный_анализ\_кандидат - Analog.xlsb"

Set objFile = objFSO.GetFile(an)
objFile.Attributes = 0
Set objFile = objFSO.GetFile(an_1)
objFile.Attributes = 0

' copy to -1
objFSO.CopyFile an, an_1, True
' copy from  candidate

objFSO.CopyFile cand_path, an, True

Set objFile = objFSO.GetFile(an)
objFile.Attributes = 1
Set objFile = objFSO.GetFile(an_1)
objFile.Attributes = 1

Set objFile = Nothing
Set objFSO  = Nothing

if not Err then
	Wscript.Echo "All Ok"
else 
	Wscript.Echo "Error:" & Chr(13) &Chr(13) & Err.Description
end if
