Attribute VB_Name = "Module2"
Option Explicit

Sub Макрос1()
Attribute Макрос1.VB_ProcData.VB_Invoke_Func = " \n14"

    If Application.ReferenceStyle = xlA1 Then
        Application.ReferenceStyle = xlR1C1
    Else
        Application.ReferenceStyle = xlA1
    End If
    
End Sub
