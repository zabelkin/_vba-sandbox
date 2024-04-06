
For i = 2 To 8
	pos = InStr(1, cell.Value, Cells(2, i).Value)
	If pos > 0 Then cell.Characters(pos, Len(Cells(2, i).Value)).Font.FontStyle = "Bold"
Next i