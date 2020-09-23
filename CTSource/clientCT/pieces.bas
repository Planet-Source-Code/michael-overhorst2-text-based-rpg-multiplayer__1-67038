Attribute VB_Name = "pieces"
Private mat(0 To 3, 0 To 1) As Integer
Public pname As Integer
Public direction As Integer
'this places the piece at its initial location
'each piece is represented by a 4x2 matrix
'4 blocks per piece each with an x and y coordinate

Public Function makePiece()
    'randomly get a piece
    Randomize
    pname = Int(Rnd * 7)
    direction = 0
    'piece 0    00
    '            00
    If pname = 0 Then
        mat(0, 0) = 3
        mat(0, 1) = 0
        mat(1, 0) = 5
        mat(1, 1) = 1
        mat(2, 0) = 4
        mat(2, 1) = 1
        mat(3, 0) = 4
        mat(3, 1) = 0
    End If
    
    'piece 1     00
    '           00
    If pname = 1 Then
        mat(0, 0) = 3
        mat(0, 1) = 1
        mat(1, 0) = 5
        mat(1, 1) = 0
        mat(2, 0) = 4
        mat(2, 1) = 1
        mat(3, 0) = 4
        mat(3, 1) = 0
    End If
    
    'piece 2    0000
    '
    If pname = 2 Then
        mat(0, 0) = 3
        mat(0, 1) = 0
        mat(1, 0) = 6
        mat(1, 1) = 0
        mat(2, 0) = 4
        mat(2, 1) = 0
        mat(3, 0) = 5
        mat(3, 1) = 0
    End If


    'piece 3    0
    '           000
    
    If pname = 3 Then
        mat(0, 0) = 3
        mat(0, 1) = 0
        mat(1, 0) = 5
        mat(1, 1) = 1
        mat(2, 0) = 4
        mat(2, 1) = 1
        mat(3, 0) = 3
        mat(3, 1) = 1
    End If
    
    'piece 4      0
    '           000
    If pname = 4 Then
        mat(0, 0) = 3
        mat(0, 1) = 1
        mat(1, 0) = 5
        mat(1, 1) = 1
        mat(2, 0) = 4
        mat(2, 1) = 1
        mat(3, 0) = 5
        mat(3, 1) = 0
    End If

    'piece 5     0
    '           000
    If pname = 5 Then
        mat(0, 0) = 3
        mat(0, 1) = 1
        mat(1, 0) = 5
        mat(1, 1) = 1
        mat(2, 0) = 4
        mat(2, 1) = 1
        mat(3, 0) = 4
        mat(3, 1) = 0
    End If
    'piece 6    00
    '           00
    
    If pname = 6 Then
        mat(0, 0) = 3
        mat(0, 1) = 0
        mat(1, 0) = 4
        mat(1, 1) = 1
        mat(2, 0) = 3
        mat(2, 1) = 1
        mat(3, 0) = 4
        mat(3, 1) = 0
    End If
    makePiece = mat
    
End Function


Public Function getname() As Integer
    'returns name of piece. name is represented by a number
    'there are 7 distinct pieces 0 - 6
    getname = pname
End Function
