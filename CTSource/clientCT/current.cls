VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "curr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'piece locations represented in this matrix
'4 blocks with x and y coords
Private piecemat(0 To 3, 0 To 1) As Integer
Private Const Xloc = 0
Private Const Yloc = 1
Public direction
Public ptype As Integer

'moves all x coords + 1
Public Sub moveRight()
For a = 0 To 3
    piecemat(a, Xloc) = piecemat(a, Xloc) + 1
Next a
End Sub
'...
Public Sub moveLeft()
For a = 0 To 3
    piecemat(a, Xloc) = piecemat(a, Xloc) - 1
Next a
End Sub
'...
Public Sub moveDown()
For a = 0 To 3
    piecemat(a, Yloc) = piecemat(a, Yloc) + 1
Next a
End Sub

'this will rotate the piece according to its type
'and current direction
'i keep the pieces so that
'block 0 is the leftmost
'block 1 is the rightmost
'block 2 is the lowest
'block 3 i don't give a crap about.
Public Sub rotate()
    If ptype = 0 And piecemat(0, 0) <= 7 Then
        If direction = 0 Or direction = 2 Then
            piecemat(0, 0) = piecemat(0, 0)
            piecemat(0, 1) = piecemat(0, 1)
            piecemat(1, 0) = piecemat(1, 0) - 1
            piecemat(1, 1) = piecemat(1, 1) - 1
            piecemat(2, 0) = piecemat(2, 0) - 1
            piecemat(2, 1) = piecemat(2, 1)
            piecemat(3, 0) = piecemat(3, 0)
            piecemat(3, 1) = piecemat(3, 1) - 1
        Else
            piecemat(0, 0) = piecemat(0, 0)
            piecemat(0, 1) = piecemat(0, 1)
            piecemat(1, 0) = piecemat(1, 0) + 1
            piecemat(1, 1) = piecemat(1, 1) + 1
            piecemat(2, 0) = piecemat(2, 0) + 1
            piecemat(2, 1) = piecemat(2, 1)
            piecemat(3, 0) = piecemat(3, 0)
            piecemat(3, 1) = piecemat(3, 1) + 1
        End If
        direction = (direction + 1) Mod 4
    End If
    If ptype = 1 And piecemat(0, 0) <= 7 Then
        If direction = 0 Or direction = 2 Then
            piecemat(0, 0) = piecemat(0, 0)
            piecemat(0, 1) = piecemat(0, 1) - 1
            piecemat(1, 0) = piecemat(1, 0) - 1
            piecemat(1, 1) = piecemat(1, 1)
            piecemat(2, 0) = piecemat(2, 0)
            piecemat(2, 1) = piecemat(2, 1)
            piecemat(3, 0) = piecemat(3, 0) - 1
            piecemat(3, 1) = piecemat(3, 1) - 1
        Else
            piecemat(0, 0) = piecemat(0, 0)
            piecemat(0, 1) = piecemat(0, 1) + 1
            piecemat(1, 0) = piecemat(1, 0) + 1
            piecemat(1, 1) = piecemat(1, 1)
            piecemat(2, 0) = piecemat(2, 0)
            piecemat(2, 1) = piecemat(2, 1)
            piecemat(3, 0) = piecemat(3, 0) + 1
            piecemat(3, 1) = piecemat(3, 1) + 1
        End If
        direction = (direction + 1) Mod 4
    End If
    If ptype = 2 And piecemat(0, 0) <= 6 Then
        If direction = 0 Or direction = 2 Then
            piecemat(0, 0) = piecemat(0, 0)
            piecemat(0, 1) = piecemat(0, 1) - 1
            piecemat(1, 0) = piecemat(1, 0) - 3
            piecemat(1, 1) = piecemat(1, 1) - 2
            piecemat(2, 0) = piecemat(2, 0) - 1
            piecemat(2, 1) = piecemat(2, 1)
            piecemat(3, 0) = piecemat(3, 0) - 2
            piecemat(3, 1) = piecemat(3, 1) - 3
        Else
            piecemat(0, 0) = piecemat(0, 0)
            piecemat(0, 1) = piecemat(0, 1) + 1
            piecemat(1, 0) = piecemat(1, 0) + 3
            piecemat(1, 1) = piecemat(1, 1) + 2
            piecemat(2, 0) = piecemat(2, 0) + 1
            piecemat(2, 1) = piecemat(2, 1)
            piecemat(3, 0) = piecemat(3, 0) + 2
            piecemat(3, 1) = piecemat(3, 1) + 3
        End If
        direction = (direction + 1) Mod 4
    End If
    If ptype = 3 And piecemat(0, 0) <= 7 Then
        If direction = 3 Then
            piecemat(0, 0) = piecemat(0, 0)
            piecemat(0, 1) = piecemat(0, 1) - 1
            piecemat(1, 0) = piecemat(1, 0) + 1
            piecemat(1, 1) = piecemat(1, 1) + 1
            piecemat(2, 0) = piecemat(2, 0)
            piecemat(2, 1) = piecemat(2, 1)
            piecemat(3, 0) = piecemat(3, 0) - 1
            piecemat(3, 1) = piecemat(3, 1) + 2
        ElseIf direction = 0 Then
            piecemat(0, 0) = piecemat(0, 0)
            piecemat(0, 1) = piecemat(0, 1)
            piecemat(1, 0) = piecemat(1, 0) - 1
            piecemat(1, 1) = piecemat(1, 1) - 2
            piecemat(2, 0) = piecemat(2, 0) - 1
            piecemat(2, 1) = piecemat(2, 1)
            piecemat(3, 0) = piecemat(3, 0)
            piecemat(3, 1) = piecemat(3, 1) - 2
        ElseIf direction = 1 Then
            piecemat(0, 0) = piecemat(0, 0)
            piecemat(0, 1) = piecemat(0, 1)
            piecemat(1, 0) = piecemat(1, 0) + 1
            piecemat(1, 1) = piecemat(1, 1) + 1
            piecemat(2, 0) = piecemat(2, 0) + 2
            piecemat(2, 1) = piecemat(2, 1)
            piecemat(3, 0) = piecemat(3, 0) + 1
            piecemat(3, 1) = piecemat(3, 1) + 1
        ElseIf direction = 2 Then
            piecemat(0, 0) = piecemat(0, 0)
            piecemat(0, 1) = piecemat(0, 1) + 1
            piecemat(1, 0) = piecemat(1, 0) - 1
            piecemat(1, 1) = piecemat(1, 1)
            piecemat(2, 0) = piecemat(2, 0) - 1
            piecemat(2, 1) = piecemat(2, 1)
            piecemat(3, 0) = piecemat(3, 0)
            piecemat(3, 1) = piecemat(3, 1) - 1
        End If
        direction = (direction + 1) Mod 4
    End If
    
    
    If ptype = 4 And piecemat(0, 0) <= 7 Then
        If direction = 0 Then
            piecemat(0, 0) = piecemat(0, 0)
            piecemat(0, 1) = piecemat(0, 1) - 1
            piecemat(1, 0) = piecemat(1, 0) - 1
            piecemat(1, 1) = piecemat(1, 1)
            piecemat(2, 0) = piecemat(2, 0) - 1
            piecemat(2, 1) = piecemat(2, 1)
            piecemat(3, 0) = piecemat(3, 0) - 2
            piecemat(3, 1) = piecemat(3, 1) - 1
        ElseIf direction = 1 Then
            piecemat(0, 0) = piecemat(0, 0)
            piecemat(0, 1) = piecemat(0, 1)
            piecemat(1, 0) = piecemat(1, 0) + 1
            piecemat(1, 1) = piecemat(1, 1) - 1
            piecemat(2, 0) = piecemat(2, 0)
            piecemat(2, 1) = piecemat(2, 1)
            piecemat(3, 0) = piecemat(3, 0) + 1
            piecemat(3, 1) = piecemat(3, 1) + 1
        ElseIf direction = 2 Then
            piecemat(0, 0) = piecemat(0, 0)
            piecemat(0, 1) = piecemat(0, 1) - 1
            piecemat(1, 0) = piecemat(1, 0) - 1
            piecemat(1, 1) = piecemat(1, 1)
            piecemat(2, 0) = piecemat(2, 0) + 1
            piecemat(2, 1) = piecemat(2, 1)
            piecemat(3, 0) = piecemat(3, 0)
            piecemat(3, 1) = piecemat(3, 1) - 1
        ElseIf direction = 3 Then
            piecemat(0, 0) = piecemat(0, 0)
            piecemat(0, 1) = piecemat(0, 1) + 2
            piecemat(1, 0) = piecemat(1, 0) + 1
            piecemat(1, 1) = piecemat(1, 1) + 1
            piecemat(2, 0) = piecemat(2, 0)
            piecemat(2, 1) = piecemat(2, 1)
            piecemat(3, 0) = piecemat(3, 0) + 1
            piecemat(3, 1) = piecemat(3, 1) + 1
        End If
        direction = (direction + 1) Mod 4
    End If
    
    If ptype = 5 And piecemat(0, 0) <= 7 Then
        If direction = 0 Then
            piecemat(0, 0) = piecemat(0, 0)
            piecemat(0, 1) = piecemat(0, 1) - 1
            piecemat(1, 0) = piecemat(1, 0) - 1
            piecemat(1, 1) = piecemat(1, 1) - 1
            piecemat(2, 0) = piecemat(2, 0) - 1
            piecemat(2, 1) = piecemat(2, 1)
            piecemat(3, 0) = piecemat(3, 0) - 1
            piecemat(3, 1) = piecemat(3, 1) - 1
        ElseIf direction = 1 Then
            piecemat(0, 0) = piecemat(0, 0)
            piecemat(0, 1) = piecemat(0, 1)
            piecemat(1, 0) = piecemat(1, 0) + 1
            piecemat(1, 1) = piecemat(1, 1)
            piecemat(2, 0) = piecemat(2, 0) + 1
            piecemat(2, 1) = piecemat(2, 1)
            piecemat(3, 0) = piecemat(3, 0) + 1
            piecemat(3, 1) = piecemat(3, 1) + 1
        ElseIf direction = 2 Then
            piecemat(0, 0) = piecemat(0, 0)
            piecemat(0, 1) = piecemat(0, 1)
            piecemat(1, 0) = piecemat(1, 0) - 1
            piecemat(1, 1) = piecemat(1, 1)
            piecemat(2, 0) = piecemat(2, 0)
            piecemat(2, 1) = piecemat(2, 1)
            piecemat(3, 0) = piecemat(3, 0)
            piecemat(3, 1) = piecemat(3, 1) - 1
        ElseIf direction = 3 Then
            piecemat(0, 0) = piecemat(0, 0)
            piecemat(0, 1) = piecemat(0, 1) + 1
            piecemat(1, 0) = piecemat(1, 0) + 1
            piecemat(1, 1) = piecemat(1, 1) + 1
            piecemat(2, 0) = piecemat(2, 0)
            piecemat(2, 1) = piecemat(2, 1)
            piecemat(3, 0) = piecemat(3, 0)
            piecemat(3, 1) = piecemat(3, 1) + 1
        End If
        direction = (direction + 1) Mod 4
    End If

End Sub


Public Function returnPiece()
'returns current locations of all 4 blocks in the piece
returnPiece = piecemat
End Function

Private Sub Class_Initialize()
'set up some initail variables
    mat = makePiece
    For a = 0 To 3
        For b = 0 To 1
            piecemat(a, b) = mat(a, b)
        Next b
    Next a
    ptype = getname
End Sub
