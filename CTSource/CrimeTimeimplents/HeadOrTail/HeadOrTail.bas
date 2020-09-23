Attribute VB_Name = "HeadOrTail"
Option Explicit

'-------------------------------------------------------'
'                     Head Or Tail Game                 '
'                     Made By Dutchbull                 '
'-------------------------------------------------------'
Public Sub HOTGen()

Dim HOTGen As String

    HOTGen = Int(Rnd * 20)
    If HOTGen <= 10 Then
        Form1.lblHOT.Caption = "Head"
    End If
    If HOTGen >= 10 Then
        Form1.lblHOT.Caption = "Tail"
    End If

End Sub

Public Sub HOTHead()

    If Form1.lblmoney.Caption <= 49 Then
        MsgBox "Not Enough Money"
    Else
        HOTGen
        Form1.lblmoney.Caption = Form1.lblmoney.Caption - 50
        If Form1.lblHOT.Caption = "Head" Then
            Form1.lblHOT2.Caption = "You Won"
            Form1.lblmoney.Caption = Form1.lblmoney.Caption + 100
        End If
        If Form1.lblHOT.Caption = "Tail" Then
            Form1.lblHOT2.Caption = "You Lost"
        End If
    End If

End Sub

Public Sub HOTTail()

    If Form1.lblmoney.Caption <= 49 Then
        MsgBox "Not Enough Money"
    Else
        HOTGen
        Form1.lblmoney.Caption = Form1.lblmoney.Caption - 50
        If Form1.lblHOT.Caption = "Tail" Then
            Form1.lblHOT2.Caption = "You Won"
            Form1.lblmoney.Caption = Form1.lblmoney.Caption + 100
        End If
        If Form1.lblHOT.Caption = "Head" Then
            Form1.lblHOT2.Caption = "You Lost"
        End If
    End If

End Sub
