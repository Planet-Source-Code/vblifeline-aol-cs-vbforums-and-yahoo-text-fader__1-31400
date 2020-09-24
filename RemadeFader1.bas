Attribute VB_Name = "Module1"
Public cColorA
Public tColor(1000000)
Public BGcolor
Public FgColor
Public BdColor
Public MdBgColor
Public MdFgColor
Public TxBgColor
Public TxFgColor
Public CbBgColor
Public cHTML
Public Function MyMSG(caption, text)
Beep
Form5.Show
Form5.Movea.caption = caption
Form5.Label1.caption = text
End Function
Public Function CustomColor()
cColorA = ""
Form2.Show
Do
DoEvents
If cColorA = "N/A" Then
CustomColor = "N/A"
Exit Function
End If
Loop Until cColorA <> ""
CustomColor = cColorA
End Function
Public Function Fade_Text(tText, Color As Variant, HTML As Boolean)
Dim red(90000), Green(90000), Blue(90000), ntext
'Can Delete This Just Used For Something Else
cHTML = HTML
'End of delete
tColor(0) = 0
X = -1
Do
DoEvents
X = X + 1
Loop Until Color(X) = ""
DoEvents
mtemp = Round(nTemp, 0) - 1
Do
mtemp = mtemp + 1
DoEvents
Loop Until mtemp * (X - 1) > Len(tText)
ncount = mtemp
For Y = 0 To X - 2
DoEvents
red(0) = 0
Green(0) = 0
Blue(0) = 0
ntext = Space(ncount)
color1 = Color(Y)
color2 = Color(Y + 1)
atemp = Hex(color1)
btemp = Hex(color2)
If Len(atemp) = 4 Then
atemp = "00" & atemp
End If
If Len(atemp) = 2 Then
atemp = "0000" & atemp
End If
If Len(btemp) = 4 Then
btemp = "00" & btemp
End If
If Len(btemp) = 2 Then
btemp = "0000" & btemp
End If
'code for scrolling blue
If Val("&h" & (Left(atemp, 2))) = Val("&h" & (Left(btemp, 2))) Then
For X = 1 To Len(ntext)
DoEvents
Blue(X) = Left(atemp, 2)
Next
Blue(0) = Blue(0) + Len(ntext)
Else
For X = Val("&h" & (Left(atemp, 2))) To Val("&h" & (Left(btemp, 2))) Step (Val("&h" & (Left(btemp, 2))) - Val("&h" & (Left(atemp, 2)))) / Len(ntext)
DoEvents
Blue(0) = Blue(0) + 1
nTemp = Hex(X)
If Len(nTemp) = 1 Then
nTemp = "0" & nTemp
End If
Blue(Blue(0)) = nTemp
Next
DoEvents
End If
'code for scrooling green
If Val("&h" & (Right(Left(atemp, 4), 2))) = Val("&h" & (Right(Left(btemp, 4), 2))) Then
For X = 1 To Len(ntext)
DoEvents
Green(X) = Right(Left(atemp, 4), 2)
Next
Green(0) = Green(0) + Len(ntext)
Else
DoEvents
For X = Val("&h" & (Right(Left(atemp, 4), 2))) To Val("&h" & (Right(Left(btemp, 4), 2))) Step (Val("&h" & (Right(Left(btemp, 4), 2))) - Val("&h" & (Right(Left(atemp, 4), 2)))) / Len(ntext)
DoEvents
Green(0) = Green(0) + 1
nTemp = Hex(X)
If Len(nTemp) = 1 Then
nTemp = "0" & nTemp
End If
Green(Green(0)) = nTemp
Next
End If
'code for scrooling red
If Val("&h" & (Right(atemp, 2))) = Val("&h" & (Right(btemp, 2))) Then
DoEvents
For X = 1 To Len(ntext)
DoEvents
red(X) = Right(atemp, 2)
Next
red(0) = red(0) + Len(ntext)
Else
For X = Val("&h" & (Right(atemp, 2))) To Val("&h" & (Right(btemp, 2))) Step (Val("&h" & (Right(btemp, 2))) - Val("&h" & (Right(atemp, 2)))) / Len(ntext)
DoEvents
red(0) = red(0) + 1
nTemp = Hex(X)
If Len(nTemp) = 1 Then
nTemp = "0" & nTemp
End If
red(red(0)) = nTemp
Next
End If
c = 0
DoEvents
For X = tColor(0) + 1 To tColor(0) + Len(ntext)
DoEvents
c = c + 1
tColor(0) = tColor(0) + 1
If HTML = False Then
tColor(X) = Blue(c) & Green(c) & red(c)
Else
tColor(X) = red(c) & Green(c) & Blue(c)
End If
DoEvents
Next
DoEvents
Next
DoEvents
End Function
Public Sub ReloadInterface()
On Error Resume Next
Dim ntempb
ntempb = ""
Set a54339 = CreateObject("WScript.Shell")
X = a54339.RegRead("HKEY_LOCAL_MACHINE\software\leoinc\fader\there")
If X = "leo" Then
Set a11313 = CreateObject("WScript.Shell")
nTemp = a11313.RegRead("HKEY_LOCAL_MACHINE\software\leoinc\fader\info")
varray = Split(nTemp, ":")
changecolor varray(0), varray(1), varray(2), varray(3), varray(4), varray(5), varray(6), varray(7)
Else
changecolor &HC07800, vbWhite, vbBlack, vbWhite, vbBlack, vbWhite, vbBlack, vbWhite
End If
Exit Sub
End Sub
Public Function changecolor(nBgColor, nFgColor, nBdColor, nMdBgColor, nMdFgColor, nTxBgColor, nTxFgColor, nCbBgColor)
nTemp = nBgColor & ":" & nFgColor & ":" & nBdColor & ":" & nMdBgColor & ":" & nMdFgColor & ":" & nTxBgColor & ":" & nTxFgColor & ":" & nCbBgColor & ":"
Set a1041 = CreateObject("WScript.Shell")
a1041.RegWrite "HKEY_LOCAL_MACHINE\software\leoinc\fader\info", nTemp
Set a57128 = CreateObject("WScript.Shell")
a57128.RegWrite "HKEY_LOCAL_MACHINE\software\leoinc\fader\there", "leo"
BGcolor = nBgColor
FgColor = nFgColor
BdColor = nBdColor
MdBgColor = nMdBgColor
MdFgColor = nMdFgColor
TxBgColor = nTxBgColor
TxFgColor = nTxFgColor
CbBgColor = nCbBgColor
If nInvis = "kktrue" Then
End
End If
UpdateColors
End Function
Public Sub UpdateColors()
For Each Var In Form1.Controls
On Error Resume Next
If Var.Tag = "t" Then
Var.BackColor = TxBgColor
Var.ForeColor = TxFgColor
End If
If Var.Tag = "cb" Then
Var.BackColor = CbBgColor
End If
If Var.Tag <> "t" And Var.Tag <> "cb" And Var.Tag <> "fb" And Var.Tag <> "1" Then
Var.BackColor = BGcolor
Var.ForeColor = FgColor
End If
Next
Form1.BackColor = BGcolor
For Each Var In Form2.Controls
On Error Resume Next
If Var.Tag = "t" Then
Var.BackColor = TxBgColor
Var.ForeColor = TxFgColor
End If
If Var.Tag = "cb" Then
Var.BackColor = CbBgColor
End If
If Var.Tag <> "t" And Var.Tag <> "cb" And Var.Tag <> "fb" Then
Var.BackColor = BGcolor
Var.ForeColor = FgColor
End If
Next
Form2.BackColor = BGcolor
For Each Var In Form3.Controls
On Error Resume Next
If Var.Tag = "t" Then
Var.BackColor = TxBgColor
Var.ForeColor = TxFgColor
End If
If Var.Tag = "cb" Then
Var.BackColor = CbBgColor
End If
If Var.Tag <> "t" And Var.Tag <> "cb" And Var.Tag <> "fb" Then
Var.BackColor = BGcolor
Var.ForeColor = FgColor
End If
Next
Form3.BackColor = BGcolor
For Each Var In Form4.Controls
On Error Resume Next
If Var.Tag = "t" Then
Var.BackColor = TxBgColor
Var.ForeColor = TxFgColor
End If
If Var.Tag = "cb" Then
Var.BackColor = CbBgColor
End If
If Var.Tag <> "t" And Var.Tag <> "cb" And Var.Tag <> "fb" Then
Var.FillColor = BGcolor
Var.BackColor = BGcolor
Var.ForeColor = FgColor
End If
Next
Form4.BackColor = BGcolor
For Each Var In Form5.Controls
On Error Resume Next
If Var.Tag = "t" Then
Var.BackColor = TxBgColor
Var.ForeColor = TxFgColor
End If
If Var.Tag = "cb" Then
Var.BackColor = CbBgColor
End If
If Var.Tag <> "t" And Var.Tag <> "cb" And Var.Tag <> "fb" Then
Var.FillColor = BGcolor
Var.BackColor = BGcolor
Var.ForeColor = FgColor
End If
Next
Form5.BackColor = BGcolor
For Each Var In Form6.Controls
On Error Resume Next
If Var.Tag = "t" Then
Var.BackColor = TxBgColor
Var.ForeColor = TxFgColor
End If
If Var.Tag = "cb" Then
Var.BackColor = CbBgColor
End If
If Var.Tag <> "t" And Var.Tag <> "cb" And Var.Tag <> "fb" Then
Var.FillColor = BGcolor
Var.BackColor = BGcolor
Var.ForeColor = FgColor
End If
Next
Form6.BackColor = BGcolor
End Sub
Public Function ErrorHandler()
Beep
Form6.Show
Form6.Label3.caption = "Details>"
Form6.Height = 1530
Form6.tt.caption = "Error:" & Err.Number & " has been created"
Form6.Label5 = Err.Description
Err.Clear
End Function

