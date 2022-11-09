Attribute VB_Name = "mdlTopMost"
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Public Sub TopMost(Form As Form)
    SetWindowPos Form.hWnd, -1, 0, 0, 0, 0, &H2 Or &H1 Or &H10
End Sub
