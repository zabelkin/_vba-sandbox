Attribute VB_Name = "sys"
Option Explicit


' before processing
Public Sub SysFunctionOff()
    
    Application.ScreenUpdating = False				' screen
    Application.Calculation = xlCalculationManual	' Calculation
    Application.EnableEvents = False				' events, including custom calls for the sheets
    ActiveSheet.DisplayPageBreaks = False			' just for clearer display
    Application.DisplayAlerts = False				' ok for most questions like sheets deletion
    
End Sub


' after processing
Public Sub SysFunctionOn()
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    ActiveSheet.DisplayPageBreaks = True
    Application.DisplayAlerts = True
    
End Sub
