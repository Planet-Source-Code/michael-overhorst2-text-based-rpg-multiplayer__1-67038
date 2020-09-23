Attribute VB_Name = "raadhetgetal"
'-------------------------------------------------------'
'                    Raad Het Getal Game                '
'                     Made By Dutchbull                 '
'    This Module Is Made For Crime Time Implenting!     '
'-------------------------------------------------------'

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


Public Sub GetalGen()
Dim Getal As String
Getal = Int(Rnd * 10)
Form1.lblgetal.Caption = Getal
End Sub
