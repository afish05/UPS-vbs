'UPSmock210.vbs
'Author: Adam Fisher
'Date: 08/18/2016
'This script will modify the 210 EDI file for UPS for re-dropping.

'NOTES:
'Make your edits, save the file as UPS.txt to the Desktop.

Const ForReading = 1
Const ForWriting = 2
Dim ctlmax, ctlmin

'Generate a random number for Interchange and Group
Randomize
ctlmax = 999
ctlmin = 100
CTLnum = Int((ctlmax-ctlmin+1)*Rnd+ctlmin)

Set objNetwork = CreateObject("Wscript.Network")
	Username = objNetwork.UserName

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.OpenTextFile("C:\Users\"&Username&"\Desktop\UPS.txt", ForReading)

Set Lines = CreateObject("VBScript.RegExp")
	Lines.Pattern = "^\w{2,3}\*"
	Lines.Global = True
	Lines.Multiline = True

Set ISA = CreateObject("VBScript.RegExp")
	ISA.Pattern = "\*\d{8}[0-9]\*0\*P"
	ISA.Global = False
	ISA.Multiline = True

Set GS = CreateObject("VBScript.RegExp")
	GS.Pattern = "\*\d{2}[0-9]\*X\*"
	GS.Global = False
	GS.Multiline = True

Set SE = CreateObject("VBScript.RegExp")
	SE.Pattern = "^SE\*.*\*"
	SE.Global = False
	SE.Multiline = True

Set GE = CreateObject("VBScript.RegExp")
	GE.Pattern = "^GE\*.*"
	GE.Global = False
	GE.Multiline = True

Set IEA = CreateObject("VBScript.RegExp")
	IEA.Pattern = "^IEA\*.*"
	IEA.Global = False
	IEA.Multiline = True

strText = objFile.ReadAll

'Receivers
DOD = Instr(strText,"*02*UPSN           *ZZ*SPSPTDOD       ")
DODI = Instr(strText,"*02*UPSN           *ZZ*SPSUPSDODINT   ")
RAY = Instr(strText,"*02*UPSN           *ZZ*SPSPTRAY       ")

'Error check
If RAY+DOD+DODI <> 32 then
	msgBox "FAILURE. The receivers ISA ID is not coded in this script. Please add it to the Receivers list."
	wscript.quit
End If

'Get the LineCount
Set LineCount = Lines.Execute(strText)
	For each match in LineCount
	Next
ISACount = LineCount.count

'Set the SE Count
SEcount = ISACount - 4

'Replace Hex 0d
strNewText = Replace(strText,chr(013),"")

'Replace ISA IDs and GS IDs
If DOD = 32 then
	strNewText1 = Replace(strNewText, "02*UPSN           *ZZ*SPSPTDOD       ", "02*UPSNSPS        *12*8004171844     ",1,1)
	strNewText2 = Replace(strNewText1, "GS*IM*UPSN*SPSPTDOD", "GS*IM*UPSNSPSDOD*8004171844",1,1)
End If

If DODI = 32 then
	strNewText1 = Replace(strNewText, "02*UPSN           *ZZ*SPSUPSDODINT   ", "02*UPSRSPS        *12*8004171844     ",1,1)
	strNewText2 = Replace(strNewText1, "GS*IM*UPSN*SPSUPSDODINT", "GS*IM*UPSRSPSDODINT*8004171844",1,1)
End If

If RAY = 32 then
	strNewText1 = Replace(strNewText, "02*UPSN           *ZZ*SPSPTRAY       ", "02*UPSNSPS        *12*8004171844     ",1,1)
	strNewText2 = Replace(strNewText1, "GS*IM*UPSN*SPSPTRAY", "GS*IM*UPSNSPSRAY*8004171844",1,1)
End If

'Replace ISA11
strNewText3 = Replace(strNewText2,"]","U",1,1)
'Replace }
strNewText4 = Replace(strNewText3, "}", ">")
'Set the ISA13
strNewText5 = ISA.Replace(strNewText4, "*000000"&CTLnum&"*0*P")
'Set the GS06
strNewText6 = GS.Replace(strNewText5, "*"&CTLnum&"*X*")
'Set the SE02 Count
strNewText7 = SE.Replace(strNewText6, "SE*"&SEcount&"*")
'Set the GE
strNewText8 = GE.Replace(strNewText7, "GE*1*"&CTLnum)
'Set the IEA
strNewText9 = IEA.Replace(strNewText8,"IEA*1*000000"&CTLnum)

Set objFile = objFSO.CreateTextFile("C:\Users\"&Username&"\Desktop\UPS-output.txt", ForWriting)

objFile.Write strNewText9

objFile.Close

MsgBox "SUCCESS. Created mocked file named UPS-output.txt on your Desktop"

wscript.quit