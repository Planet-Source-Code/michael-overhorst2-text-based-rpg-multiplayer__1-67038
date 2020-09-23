Attribute VB_Name = "Higherlowergame"
'-------------------------------------------------------'
'                     Higher Lower Game                 '
'                     Made By Dutchbull                 '
'    This Module Is Made For Crime Time Implenting!     '
'-------------------------------------------------------'



Public Sub Lager()
On Error Resume Next
Dim Newnumber2 As String
Dim oldnumber2 As String
If Form1.lblmoney.Caption <= 49 Then
MsgBox "not enough money"
Else
Form1.lblmoney.Caption = Form1.lblmoney.Caption - 50
Newnumber2 = Int(Rnd * 50)
oldnumber2 = Int(Rnd * 50)
Form1.lblnewnumber.Caption = Newnumber2
If Newnumber2 <= Form1.lbloldnumber.Caption Then
Form1.lblhl.Caption = "You Won"
Form1.lblmoney.Caption = Form1.lblmoney.Caption + 100
Else
Form1.lblhl.Caption = "You Lost"
End If
Form1.lbloldnumber.Caption = oldnumber2
End If
End Sub




Public Sub Hoger()
On Error Resume Next
Dim Newnumber As String
Dim oldnumber As String
If Form1.lblmoney.Caption <= 49 Then
MsgBox "not enough money"
Else
Form1.lblmoney.Caption = Form1.lblmoney.Caption - 50
Newnumber = Int(Rnd * 50)
oldnumber = Int(Rnd * 50)
Form1.lblnewnumber.Caption = Newnumber

If Newnumber >= Form1.lbloldnumber.Caption Then
Form1.lblhl.Caption = "You Won"
Form1.lblmoney.Caption = Form1.lblmoney.Caption + 100
Else
Form1.lblhl.Caption = "You Lost"
End If
Form1.lbloldnumber.Caption = oldnumber
End If
End Sub

Public Sub GenNumbers()
Dim Newnumber As String
Dim oldnumber As String
Newnumber = Int(Rnd * 50)
oldnumber = Int(Rnd * 50)
Form1.lbloldnumber.Caption = oldnumber
Form1.lblnewnumber.Caption = Newnumber
Form1.lblhl.Caption = "Higher Of Lower"
End Sub
