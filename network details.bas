Sub GetUserName_Environ()
    Dim ObjWshNw As Object
    Set ObjWshNw = CreateObject("WScript.Network")
    
    MsgBox ObjWshNw.UserName
    MsgBox ObjWshNw.ComputerName
    MsgBox ObjWshNw.UserDomain
End Sub