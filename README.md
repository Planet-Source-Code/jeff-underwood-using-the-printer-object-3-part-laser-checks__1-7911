<div align="center">

## Using the Printer Object 3 Part Laser Checks


</div>

### Description

This module will show how to use the printer object

programmatically, in this case to print 3 part laser

checks but anyone can use this on a laser printer with

plain paper. I think it is about compact as it can be

there are some variables that are not used yet. Notes

on their purpose are documented.

This shows how to use Printer.TextHeight and

Printer.TextWidth for something useful.

Please rate my code if you find it useful. ;)
 
### More Info
 
None required.

Just place a button on a form, add this module to the

project and in the OnClick event of the button call

DoCheckDemo.

This may not work with all printer drivers, but does work fine with

laserjet II and 4 drivers.

None know, it might print 4 pages or more with some

printer drivers, in this case the values in the

Init3PartLaserChecks will need modification.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Jeff Underwood](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/jeff-underwood.md)
**Level**          |Advanced
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Files/ File Controls/ Input/ Output](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/files-file-controls-input-output__1-3.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/jeff-underwood-using-the-printer-object-3-part-laser-checks__1-7911/archive/master.zip)





### Source Code

```
'This module was made for printing data on preprinted 3 part laser checks
'With the actual check at the top, and 2 stub sections below
'It is designed specifically for McBee Form LTM101-1R (I believe this is the form #)
' J13,20000410,01018020602001,000001012900 are also #'s that are on the side I'm sure any
'
'I made this to get the data from an array so you can use the code and learn about the Printer Object
'without a data base.  DoCheckDemo will print a random check for you just place a button on a form
'add this module and put DoCheckDemo in the OnClick Event of the button
'
'It will print 80 line items on these checks , to use this on any other form should be as easy as
'modifying the values in Init3PartLaserChecks in the event you want to use them with a current
'form that may have the check in the middle or on the bottom.
'
'I have taken great care to name these variables descriptive therefore they are long, but descriptive
'Also there are very few examples of programs showing programatic use of the printer.object
'so there are things in here that are not necessarily the best (or easiest) way, but
'it shows FUNCTIONAL use of the Printer.TextHeight and Printer.TextWidth
'
'If you use this in a program please let me know, at least say thanks and let me see what you did with it
'Also, if you know how I could have done this in Crystal Reports E-mail me w/ info
'
'
Public StubItems(80, 4) As String       'Up to 50 items per check stub each item can have 5 columns
Public StubHeader(5, 1) As Variant     '
Public CheckItems(8)  '0=PayStr 1=ChkDate 2=ChkAmt 3-PayName 4=PayAdd1
                   ' 5=PatAdd2,6=CityStZip, 7=Attn (Optional If Present goes after PayName)
Public StubItemCount As Integer        'The number of invoices that are paid on the check (# line items on stubs)
Public StubHeaderFields As Byte
Public MaxStubLines As Byte          'Maximum # of lines to print on each stub
Public PayAmtString As String         'NINE THOUSAND NINE HUNDRED etc..
Public PayAmtStringX As Integer       'X,Y Location to Print NINE THOUSAND NINE HUNDRED etc..
Public PayAmtStringY As Integer
Public CheckTopY As Integer          'Top of Check 0 is fine unless the check is in a position other than the top of the page
Public EnvWinTopY As Integer         'Cordinates of where on the page the Name address should go
Public EnvWinBotY As Integer         'So they show up in the envelope window
Public EnvWinLeftX As Integer
Public EnvWinRightX As Integer
Public EnvWinFontSize As Integer
Public ChkDate As String             'Check Date
Public ChkDateX As Integer           'X,Y Location to print Check Date
Public ChkDateY As Integer
Public ChkAmt As String             '$9,999.99
Public ChkAmtX As Integer           'X,Y Location to print
Public ChkAmtY As Integer
Public StubSpacing As Integer        'Horizontal spacing of Stub Columns
Public Stub1TopY As Integer         'Top and bottom value (Y) of stub1 and 2
Public Stub2TopY As Integer         'The bottom values are not actually used yet
Public Stub1BotY As Integer          'but will be needed to make the routine dynamically size the font and
Public Stub2BotY As Integer          'change the spacing for varying #'s of line items
Public ChkStubColSpace As Integer
Public ChkStubSect1StartX As Integer
Sub Print3PartLaserChecks(StubLines As Integer)
StubItemCount = StubLines
If StubItemCount < 1 Then Exit Sub
PrintCheck
PrintStubs StubItemCount
'Printer.KillDoc
Printer.EndDoc
End Sub
Sub Init3PartLaserChecks()
MaxStubLines = 80
CheckTopY = 0
ChkDateX = 7750       'X,Y Location to print Check Date
ChkDateY = 2250
PayAmtStringX = 1250   'X,Y Location to Print NINE THOUSAND NINE HUNDRED etc..
PayAmtStringY = 2250
ChkAmtX = 9600        'X,Y Location to print "$9,999.99"
ChkAmtY = 2250
EnvWinTopY = 3000     'X,Y Locations of area of Laser check that will show in a standard window envelope
EnvWinBotY = 3900
EnvWinLeftX = 1200
EnvWinRightX = 5500
Stub1TopY = 5100      'The Top (Y) position for Stub 1 ( use a # after perforation so it doesn't print over comp name & check num)
Stub1BotY = 9800       'The Bottom (Y) position for stub 1 (not in use at the moment going to use this for making the
                     'stubs use a range of font sizes depending on the number of items so 60 or so total items
Stub2TopY = 10300     'can be paid with one check using the smallest font but checks with 15 or 20 items will use
Stub2BotY = 13900      'a more reasonable font... Right now on one of the very common layouts using a font size 6
                     'you can get around 60 items per check. This is gonna save a client about 15 checks a month
                     'because the current system can only get 20 entries on a stub then it prints a wasted check
                     'voided with remaining info on subsequent stubs. (sometimes 3 or 4 of them)
ChkStubColSpace = 1100   'Spacing between the headings and stub entries on both stubs
ChkStubSect1StartX = 250  'Sets how far in (in addition to the regular print margin!) to start printing stub headers/entries
'Define the Stub Header Fields      This is probably how anyone will need this, however by changing the array you
'                              can add something or remove say DISC AMT (Discount Amt)
StubHeader(0, 0) = "INV DATE"
StubHeader(1, 0) = "INV NUM"
StubHeader(2, 0) = "INV AMT"
StubHeader(3, 0) = "DISC AMT"
StubHeader(4, 0) = "AMT PAID"
StubHeader(0, 1) = vbLeftJustify
StubHeader(1, 1) = vbLeftJustify
StubHeader(2, 1) = vbRightJustify
StubHeader(3, 1) = vbRightJustify
StubHeader(4, 1) = vbRightJustify
StubHeaderFields = 5     'Not really needed but easier for the beginners to understand than UBound
End Sub
Sub PrintStubs(StubItemCount)
Dim StubLine As Byte
Dim StubCol As Byte
Dim ChkStubLineItemSpace As Byte
Printer.FontSize = 8
Stub1YPos = Printer.TextHeight("Z,") + Stub1TopY
Stub2YPos = Printer.TextHeight("Z,") + Stub2TopY
Printer.FontSize = 6
'Multiplying the following line by .8 just takes away some extra spacing between the lines
'to get more items on the check
ChkStubLineItemSpace = Printer.TextHeight(StubItems(0, 0)) * 0.8
PrintStubHeaders StubItemCount
For StubLine = 0 To StubItemCount - 1
   'Next line just checks to see if the line count needs to print in the left or right detail area of the stub
   'If it does then it just adds 1/2 of the width of the printing area and prints the right with the same format
   'adding the additional spacing specified by ChkStubSect1StartX (In Init routine)
   'Saved having to duplicate these in a if then else or an extra loop
   If StubLine > (MaxStubLines / 2) - 1 Then StubLineMult = StubLine - (MaxStubLines / 2) Else StubLineMult = StubLine ' This is The Left Group of Cols on the Stub
     For StubCol = 0 To StubHeaderFields - 1
        Printer.CurrentX = FormatStubLine(StubLine, StubCol)
        Printer.CurrentY = Stub1YPos + (ChkStubLineItemSpace * StubLineMult)
        Printer.Print StubItems(StubLine, StubCol)
        Printer.CurrentX = FormatStubLine(StubLine, StubCol)
        Printer.CurrentY = Stub2YPos + (ChkStubLineItemSpace * StubLineMult)
        Printer.Print StubItems(StubLine, StubCol)
     Next StubCol
Next StubLine
End Sub
Function FormatStubLine(SLine As Byte, SCol As Byte) As Integer
If SLine > (MaxStubLines / 2) - 1 Then StubSect = Printer.ScaleWidth / 2 Else StubSect = 0
'When you fill the array columns yu can specify vbRightJustify (1) or vbLeftJustify(0 default) in the array
If StubHeader(SCol, 1) = vbLeftJustify Then FormatStubLine = ChkStubSect1StartX + StubSect + (ChkStubColSpace * SCol)
If StubHeader(SCol, 1) = vbRightJustify Then
  hdrPrintStartX = ChkStubSect1StartX + StubSect + (ChkStubColSpace * SCol)
  hdrPrintWidth = Printer.TextWidth(StubHeader(SCol, 0))
  StubItemPrintWidth = Printer.TextWidth(StubItems(SLine, SCol))
  'This will Align decimal figures to print right aligned with the header above them
  FormatStubLine = hdrPrintStartX + hdrPrintWidth - StubItemPrintWidth
End If
End Function
Sub PrintStubHeaders(StubItemCount)
Printer.FontBold = True
Printer.FontUnderline = True
For Shdr = 0 To StubHeaderFields - 1
   Printer.CurrentX = ChkStubSect1StartX + (ChkStubColSpace * Shdr)
   Printer.CurrentY = Stub1TopY
   Printer.Print StubHeader(Shdr, 0)
   Printer.CurrentX = ChkStubSect1StartX + (ChkStubColSpace * Shdr)
   Printer.CurrentY = Stub2TopY
   Printer.Print StubHeader(Shdr, 0)
'Print the 2nd column header only if necessary
   If StubItemCount > (MaxStubLines / 2) - 1 Then
     Printer.CurrentX = ChkStubSect1StartX + (ChkStubColSpace * Shdr) + Printer.ScaleWidth / 2
     Printer.CurrentY = Stub1TopY
     Printer.Print StubHeader(Shdr, 0)
     Printer.CurrentX = ChkStubSect1StartX + (ChkStubColSpace * Shdr) + Printer.ScaleWidth / 2
     Printer.CurrentY = Stub2TopY
     Printer.Print StubHeader(Shdr, 0)
   End If
   'ChkStubSect1StartX = ChkStubSect1StartX + ChkStubColSpace
Next Shdr
Printer.FontBold = False
Printer.FontUnderline = False
End Sub
Sub PrintCheck()
'Dim CheckItems(8)  '0=PayStr 1=ChkDate 2=ChkAmt 3-PayName 4=PayAdd1
'                ' 5=PatAdd2,6=CityStZip, 7=Attn (Optional If Present goes after PayName)
Printer.CurrentX = PayAmtStringX
Printer.CurrentY = PayAmtStringY
Printer.Font = "Arial Narrow"
Printer.FontSize = 10
Printer.FontBold = False
Printer.Print CheckItems(0)   '"NINE THOUSAND NINE HUNDRED NINETY NINE AND 99/100 ************************"
Printer.CurrentX = ChkDateX
Printer.CurrentY = ChkDateY
Printer.Font = "Arial"
Printer.Print CheckItems(1)   '"12/31/2000"
Printer.CurrentX = ChkAmtX
Printer.CurrentY = ChkAmtY
Printer.FontSize = 12
Printer.FontBold = True
Printer.Print CheckItems(2)  '"***$9,999.99"
Printer.CurrentX = EnvWinLeftX
Printer.CurrentY = EnvWinTopY
Printer.FontBold = False
Printer.FontSize = 12
EnvWindowLineCount = 0
LineHeight = Printer.TextHeight(CheckItems(3))
Printer.Print CheckItems(3)  ' "PAYNAMEPAYNAMEPAYNAMEPAYNAME"
EnvWindowLineCount = EnvWindowLineCount + 1
If Trim(CheckItems(7)) <> "" Then
  Printer.FontBold = True
  Printer.FontUnderline = True
  Printer.CurrentX = EnvWinLeftX
  Printer.CurrentY = EnvWinTopY + (LineHeight * EnvWindowLineCount)
  Printer.Print CheckItems(7)
  EnvWindowLineCount = EnvWindowLineCount + 1
  Printer.FontBold = False
  Printer.FontUnderline = False
End If
Printer.CurrentX = EnvWinLeftX
Printer.CurrentY = EnvWinTopY + (LineHeight * EnvWindowLineCount)
Printer.Print CheckItems(4)   ' "PAYADD1PAYADD1PAYADD1PAYADD1"
EnvWindowLineCount = EnvWindowLineCount + 1
If Trim(CheckItems(5)) <> "" Then
  Printer.CurrentX = EnvWinLeftX
  Printer.CurrentY = EnvWinTopY + (LineHeight * EnvWindowLineCount)
  Printer.Print CheckItems(5)   ' "PAYADD2PAYADD2PAYADD2PAYADD2"
  EnvWindowLineCount = EnvWindowLineCount + 1
End If
Printer.CurrentX = EnvWinLeftX
Printer.CurrentY = EnvWinTopY + (LineHeight * EnvWindowLineCount)
Printer.Print CheckItems(6)   ' "CITYSTATEZIPCITYSTATEZIP"
End Sub
Sub DoCheckDemo()
'Just add a Button to a form and put DoCheckDemo in the on click event
'This will print a sample of what a check would look like you can then
'easily play with the values to line them up for your particular need
'
Init3PartLaserChecks
Randomize
'Init3PartLaserChecks
 '0=PayStr 1=ChkDate 2=ChkAmt 3=PayName 4=PayAdd1, 5=PatAdd2,6=CityStZip, 7=Attn (Optional If Present goes after PayName)
 CheckItems(0) = "Nine Thousand Nine Hundred Ninety Nine and 99/100 *******************"
 CheckItems(1) = "12/31/2000"
 CheckItems(2) = "***" + "9,999.99" + "***"
 CheckItems(3) = "John D. Doe"
 CheckItems(4) = "123 Anystreet"
' CheckItems(5) = ""
 CheckItems(6) = "Anytown, AnyState 99999-9999"
For InsLine = 0 To 79
   StubItems(InsLine, 0) = "12/31/2000"
   StubItems(InsLine, 1) = Str(Int((999999 - 999 + 1) * Rnd + 999))
   StubItems(InsLine, 2) = Format((99999.99 - 999.99 + 1) * Rnd + 99.99, "Currency")
   StubItems(InsLine, 3) = Format((99.99 - 9.99 + 1) * Rnd + 0.99, "Currency")
   StubItems(InsLine, 4) = Format(StubItems(InsLine, 2) - StubItems(InsLine, 3), "Currency")
Next InsLine
Print3PartLaserChecks Int(InsLine)
End Sub
```

