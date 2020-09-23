Attribute VB_Name = "paperscissorsstone"
Option Explicit

Public Sub ComputerGen()

Dim comphand As String

    On Error Resume Next
    comphand = Int(Rnd * 30)
    If comphand >= 0 Then
        Form1.lblcomphand.Caption = "Paper"
    End If
    If comphand >= 10 Then
        Form1.lblcomphand.Caption = "Stone"
    End If
    If comphand >= 20 Then
        Form1.lblcomphand.Caption = "Scissors"
    End If
    On Error GoTo 0
End Sub


Public Sub Paper()

    If Form1.lblmoney.Caption <= 49 Then
        MsgBox "Not enough money"
    Else
        Form1.lblmoney.Caption = Form1.lblmoney.Caption - 50
        If Form1.lblcomphand.Caption = "Scissors" Then
            Form1.lblpss.Caption = "You Lost"
        End If
        If Form1.lblcomphand.Caption = "Paper" Then
            Form1.lblpss.Caption = "Draw Game"
        End If
        With Form1
            If .lblcomphand.Caption = "Stone" Then
                .lblpss.Caption = "You Won"
                .lblmoney.Caption = .lblmoney.Caption + 100
            End If
        End With
    End If

End Sub

Public Sub Scissors()

    If Form1.lblmoney.Caption <= 49 Then
        MsgBox "Not enough money"
    Else
        Form1.lblmoney.Caption = Form1.lblmoney.Caption - 50
        If Form1.lblcomphand.Caption = "Scissors" Then
            Form1.lblpss.Caption = "Draw Game"
        End If
        With Form1
            If .lblcomphand.Caption = "Paper" Then
                .lblpss.Caption = "You Won"
                .lblmoney.Caption = .lblmoney.Caption + 200
            End If
        End With
        If Form1.lblcomphand.Caption = "Stone" Then
            Form1.lblpss.Caption = "You Lost"
        End If
    End If

End Sub

Public Sub Stone()

    If Form1.lblmoney.Caption <= 49 Then
        MsgBox "Not enough money"
    Else
        Form1.lblmoney.Caption = Form1.lblmoney.Caption - 50
        If Form1.lblcomphand.Caption = "Scissors" Then
            Form1.lblpss.Caption = "You Won"
            Form1.lblmoney.Caption = Form1.lblmoney.Caption + 100
        End If
        If Form1.lblcomphand.Caption = "Paper" Then
            Form1.lblpss.Caption = "You Lost"
        End If
        If Form1.lblcomphand.Caption = "Stone" Then
            Form1.lblpss.Caption = "Draw Game"
        End If
    End If

End Sub
