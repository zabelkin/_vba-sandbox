Call showStatus(Current, Total, "  Process Running: ")

Private Sub showStatus(Current As Integer, lastrow As Integer, Topic As String)
Dim NumberOfBars As Integer
Dim pctDone As Integer

NumberOfBars = 50
'Application.StatusBar = "[" & Space(NumberOfBars) & "]"


' Display and update Status Bar
    CurrentStatus = Int((Current / lastrow) * NumberOfBars)
    pctDone = Round(CurrentStatus / NumberOfBars * 100, 0)
    Application.StatusBar = Topic & " [" & String(CurrentStatus, "|") & _
                            Space(NumberOfBars - CurrentStatus) & "]" & _
                            " " & pctDone & "% Complete"

' Clear the Status Bar when you're done
'    If Current = Total Then Application.StatusBar = ""

End Sub