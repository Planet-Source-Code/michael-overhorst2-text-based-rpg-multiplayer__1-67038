Attribute VB_Name = "Module1"

Function FileExist(FileName As String) As Boolean
    On Error GoTo NotExist
    Close #1
    Open FileName For Input As #1
    Close #1
    FileExist = True
    Exit Function
NotExist:
    FileExist = False
End Function
