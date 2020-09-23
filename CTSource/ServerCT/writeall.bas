Attribute VB_Name = "save"
Option Explicit


'------------------------------------------'
'Save Module gemaakt door michael overhorst'
'------------------------------------------'
'lstData = de listbox                      '
'de listbox waar hij dingen in laad!       '
'LoadList save.Text, lstData               '
'------------------------------------------'
'save.text is een textbox met de path erin '
'Bijvb. c:/text.txt of .wat je wilt        '
'save.text = waar hij moet opslaan         '
'------------------------------------------'
'om op teslaan'                            '
'SaveList save.Text, lstData               '
'handig voor dingen zoals chatloggen enzo  '
'------------------------------------------'


Public Sub LoadList(sLocation As String, lstListBox As ListBox)
On Error Resume Next
Dim sCurrent As String
Dim I As Integer



lstListBox.Clear

Open sLocation For Input As #1

I = 0




Do Until EOF(1)

Line Input #1, sCurrent

lstListBox.AddItem sCurrent, I

I = I + 1


Loop

Close #1



Exit Sub

Exit Sub
End Sub

Public Sub SaveList(sLocation As String, lstListBox As ListBox)
On Error Resume Next

Dim sCurrent As String
Dim I As Integer

Open sLocation For Output As #1


I = 0

Do Until I = lstListBox.ListCount

sCurrent = lstListBox.List(I)

Print #1, sCurrent

I = I + 1


Loop


Close #1


Exit Sub

Exit Sub
End Sub
