Public Sub EnvironSetItem(environVariable As String, Optional newValue As String = "", Optional strType As String = "process")

    With CreateObject("WScript.Shell").Environment(strType) ' "process" is much faster than "user", so recommended
        .Item(environVariable) = newValue
    End With

End Sub

Public Function EnvironGetItem(ByVal environVariable As String, Optional ByVal strType As String = "process") As String
    
    With CreateObject("WScript.Shell").Environment(strType)
        EnvironGetItem = .Item(environVariable)
    End With

End Function

Sub test()

    Call EnvironSetItem("myValue", "someValue")
    Debug.Print "myValue1=" & EnvironGetItem("myValue")
    
    Call EnvironSetItem("myValue")
    Debug.Print "myValue2=" & EnvironGetItem("myValue")

End Sub