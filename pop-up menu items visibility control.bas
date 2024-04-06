Attribute VB_Name = "Module1"
Option Explicit

Public Sub доступное_меню()
    Dim ctrl As Office.CommandBarControl
    Dim arrIdx(13) As Long
    Dim idx As Variant
    Dim mnuControls As CommandBarControls
 
    ' Assign values to each element of the array - arrIdx, this can be added to as needed ;)
    arrIdx(0) = 21
    arrIdx(1) = 3181
    arrIdx(2) = 21437
    
'    arrIdx(0) = 292         ' Cell Delete
'    arrIdx(1) = 293         ' Row Delete
'    arrIdx(2) = 294         ' Column Delete
'    arrIdx(3) = 295         ' Cell Insert
'    arrIdx(4) = 27960       ' Row & Column Insert
'    arrIdx(5) = 3125        ' Clear Contents
'    arrIdx(6) = 31402       ' Cell Filter
'    arrIdx(7) = 31435       ' Cell Sort
'    arrIdx(8) = 541         ' Row Height
'    arrIdx(9) = 542         ' Column Height
'    arrIdx(10) = 883        ' Row Hide
'    arrIdx(11) = 884        ' Row Unhide
'    arrIdx(12) = 886        ' Column Hide
'    arrIdx(13) = 887        ' Column Unhide
 
    For Each idx In arrIdx
        Set mnuControls = Application.CommandBars.FindControls(ID:=idx)
        If Not mnuControls Is Nothing Then 'If no CommandBarControls were found skip the following
            For Each ctrl In mnuControls
                ctrl.Enabled = True
            Next ctrl
        End If
    Next idx
 
End Sub

Public Sub список_меню()
    Dim ctrl As Office.CommandBarControl
    Dim arrIdx(13) As Long
    Dim idx As Long
    Dim mnuControls As CommandBarControls
  
    For idx = 1 To 65000
        Set mnuControls = Application.CommandBars.FindControls(ID:=idx)
        If Not mnuControls Is Nothing Then 'If no CommandBarControls were found skip the following
        Debug.Print idx
'            For Each ctrl In mnuControls
'                ctrl.Enabled = True
'            Next ctrl
        End If
    Next idx
 
End Sub


Sub simple()
    CommandBars("Cell").Reset
    CommandBars("Row").Reset
    CommandBars("Column").Reset
End Sub

