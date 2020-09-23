Attribute VB_Name = "raadhetgetal"
Option Explicit

Public Sub GetalGen()

Dim Getal As String

    Getal = Int(Rnd * 10)
    Form1.lblgetal.Caption = Getal

End Sub

Public Sub RaadGetal()

    If Form1.lblmoney.Caption <= 49 Then
        MsgBox "not enough money"
    Else
        GetalGen
        Form1.lblmoney.Caption = Form1.lblmoney.Caption - 50
        If Form1.txtgetal.Text = Form1.lblgetal.Caption Then
            Form1.lblinfoo = "You Won"
            Form1.lblmoney.Caption = Form1.lblmoney.Caption + 300
        Else
            Form1.lblinfoo = "You Lost"
        End If
    End If

End Sub

