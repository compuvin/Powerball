Dim MyValue, pt, i, j, GrpCount, ThisGrp, countComma, posNumber, PrevWinners, isWin, RankComp, FoundRank, RankSetCount, RankZeros, RankAll, response
Dim RndAmt 'How many random numbers to choose
Dim SetsAr(4) 'sets
Dim RankSets(4) 'rank each set
Dim RankTotals(9) 'total of each ranked number
dim Grp1, Grp2, Grp3, Grp4, Grp5, GrpTemp, GrpTempAll 'groups
dim Check1, Check2, checkYN 'checks
Dim LastLotto, WinCount, WinFive 'last winning number
Dim WinningNbrs(4) 'Previous five winning numbers
dim WinPair(59,1) 'winning numbers pair with
dim WinOrig(59,1), TheThree, WinTres(59,1), PickThree 'winning numbers pair with first two
Dim PlayRepeats 'Whether to play repeats or not
dim OutputNote 'generate output file
dim filesys
set filesys=CreateObject("Scripting.FileSystemObject")

OutputNote = "Powerball file generated on " & format(date(), "MM/DD/YYYY") & " for drawing on " & format(date() + SatWed(date()), "MM/DD/YYYY") & vbcrlf

'1 through 59
'All numbers: "1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31,32,33,34,35,36,37,38,39,40,41,42,43,44,45,46,47,48,49,50,51,52,53,54,55,56,57,58,59"
Grp1 = "1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31,32,33,34,35,36,37,38,39,40,41,42,43,44,45,46,47,48,49,50,51,52,53,54,55,56,57,58,59"
Grp2 = Grp1
Grp3 = Grp1
Grp4 = Grp1
Grp5 = Grp1


PrevWinners = getfile("PreviousPB.csv")
LastLotto = mid(PrevWinners,instr(1,PrevWinners,"|")+1,instr(1,PrevWinners,",P")-instr(1,PrevWinners,"|"))
'msgbox LastLotto

'Remove last powerballs (1 right now) - Added 2015-01-15
dim LastPB(0), PBTemp
PBTemp = PrevWinners
for i = 0 to ubound(LastPB)
  LastPB(i) = int(mid(PBTemp,instr(PBTemp,"P")+1,2))
  PBTemp = replace(PBTemp,left(PBTemp,instr(PBTemp,"P")+3),"X")
  'msgbox "'" & LastPB(i) & "'"
next
for i = 0 to ubound(LastPB) 'Separated for portability :)
  Grp2 = Replace("," & Grp2 & ",", "," & LastPB(i) & ",", ",")
  if left(Grp2, 1) = "," then Grp2 = right(Grp2, len(Grp2)-1)
  if right(Grp2, 1) = "," then Grp2 = left(Grp2, len(Grp2)-1)
next

'Remove previous winners
for WinCount = 0 to 4
  WinningNbrs(WinCount) = int(left(LastLotto,instr(1,LastLotto,",")-1))
  LastLotto = replace("," & LastLotto,"," & WinningNbrs(WinCount) & ",", "")
  'msgbox WinningNbrs(WinCount)
  
  Grp2 = Replace("," & Grp2 & ",", "," & WinningNbrs(WinCount) & ",", ",")
  if left(Grp2, 1) = "," then Grp2 = right(Grp2, len(Grp2)-1)
  if right(Grp2, 1) = "," then Grp2 = left(Grp2, len(Grp2)-1)
  
  Grp3 = Grp2
  Grp4 = Grp2
  Grp5 = Grp2
next
GrpTempAll = Grp2

'Decide whether to play repeats
PlayRepeats = inputbox("Enter how many REPEAT numbers to pick:", "REPEAT Numbers", "0") 
if PlayRepeats <= 0 then
  OutputNote = OutputNote & vbcrlf & "'Repeat numbers were skipped" & vbcrlf
else
PlayRepeats = PlayRepeats - 1 'List starts at 0 so taking one off
for WinCount = 0 to 4
  OutputNote = OutputNote & vbcrlf & "'" & WinningNbrs(WinCount) & " (Count = " & (len(PrevWinners) - len(replace(replace(PrevWinners,"|",","), "," & WinningNbrs(WinCount) & ",", ","))) / (len(WinningNbrs(WinCount)) + 1) & ")" & vbcrlf
  Grp1 = WinningNbrs(WinCount)
  
  
  'Rank numbers
  for i = 0 to 59
    WinPair(i,0) = i
    WinPair(i,1) = 0
  next
  RankComp = PrevWinners
	  
  j = InStr(1,replace(RankComp, "|", ","),"," & WinningNbrs(WinCount) & ",")
  do while j > 0
    FoundRank = mid(RankComp,InStrRev(RankComp,"|",j) + 1, InStr(j,RankComp,"P") - InStrRev(RankComp,"|",j) - 2)
    'msgbox InStrRev(RankComp,"|",j) & " " & InStr(j,RankComp,"P") & vbcrlf & FoundRank
	
	'RankSets(i) = RankSets(i) & "," & FoundRank & ","
	for i = 0 to 59
	  if instr(1,"," & FoundRank & ",", "," & i & ",") and not i = WinningNbrs(WinCount) then
        WinPair(i,1) = WinPair(i,1) + 1
	  end if
    next
	RankComp = replace(RankComp, FoundRank, "0,0,0,0,0")
		
	'Look again
	j = InStr(1,replace(RankComp, "|", ","),"," & WinningNbrs(WinCount) & ",")
  loop
  for i = 0 to 4
    WinPair(WinningNbrs(i),1) = 0
  next
  
  'Sort Array
  for i = 0 to 59
    WinOrig(i,0) = WinPair(i,0)
	WinOrig(i,1) = WinPair(i,1)
  next
  SortB WinPair
    'dim testtest
  'testtest = ""
  'for i = 0 to 59
  '  testtest = testtest & WinPair(i,0) & " - " & WinPair(i,1) & vbcrlf
  'next
  'msgbox WinningNbrs(WinCount) & vbcrlf & vbcrlf & testtest
  
  
  
  

  'Remove top 3 from list
  GrpTemp = Grp2
  if PlayRepeats > 2 then
    for i = 0 to 2
	  GrpTemp = Replace("," & GrpTemp & ",", "," & WinPair(i,0) & ",", ",")
      if left(GrpTemp, 1) = "," then GrpTemp = right(GrpTemp, len(GrpTemp)-1)
      if right(GrpTemp, 1) = "," then GrpTemp = left(GrpTemp, len(GrpTemp)-1)
	  Grp3 = GrpTemp
      Grp4 = GrpTemp
      Grp5 = GrpTemp
    next
  else
    for i = 0 to PlayRepeats
	  GrpTemp = Replace("," & GrpTemp & ",", "," & WinPair(i,0) & ",", ",")
      if left(GrpTemp, 1) = "," then GrpTemp = right(GrpTemp, len(GrpTemp)-1)
      if right(GrpTemp, 1) = "," then GrpTemp = left(GrpTemp, len(GrpTemp)-1)
	  Grp3 = GrpTemp
      Grp4 = GrpTemp
      Grp5 = GrpTemp
    next
  end if

	  
  for WinFive = 0 to PlayRepeats
    if WinFive < 4 then
	  Grp2 = WinPair(WinFive,0)
	  
	  'Rank third number
	  for i = 0 to 59
	    WinTres(i,0) = i
    	WinTres(i,1) = 0
	  next
	  RankComp = PrevWinners
	
      j = InStr(1,replace(RankComp, "|", ","),"," & Grp2 & ",")
	  do while j > 0
    	FoundRank = mid(RankComp,InStrRev(RankComp,"|",j) + 1, InStr(j,RankComp,"P") - InStrRev(RankComp,"|",j) - 2)
      	'msgbox InStrRev(RankComp,"|",j) & " " & InStr(j,RankComp,"P") & vbcrlf & FoundRank
	
	  	'RankSets(i) = RankSets(i) & "," & FoundRank & ","
	  	for i = 0 to 59
	    	if instr(1,"," & FoundRank & ",", "," & i & ",") and not i = Grp2 then
          	WinTres(i,1) = WinTres(i,1) + 1
	    	end if
      	next
	  	RankComp = replace(RankComp, FoundRank, "0,0,0,0,0")
		
		'Look again
		j = InStr(1,replace(RankComp, "|", ","),"," & Grp2 & ",")
      loop
  	  for i = 0 to 4
    	WinTres(WinningNbrs(i),1) = 0
  	  next
  	  for TheThree = 0 to 59
	    if instr(1,PickThree, "," & TheThree & ",") > 0 then
		  WinTres(TheThree,1) = 0
		else
    	  WinTres(TheThree,1) = WinTres(TheThree,1) * WinOrig(TheThree,1)
		end if
  	  next
   
  
  
  
  
  
  	for i = 0 to 2
		WinTres(int(WinPair(i,0)),1) = 0
 	 next
 	 SortB WinTres
	 Grp3 = WinTres(0,0)
	 PickThree = PickThree & "," & WinTres(0,0) & ","
	 'msgbox PickThree
	 'msgbox WinningNbrs(WinCount) & vbcrlf & vbcrlf & WinPair(0,0) & " - " & WinPair(0,1) & ", " & WinTres(0,0) & " - " & WinTres(0,1) & _
 	 ' vbcrlf & WinPair(1,0) & " - " & WinPair(1,1) & ", " & WinTres(1,0) & " - " & WinTres(1,1) & _
	 ' vbcrlf & WinPair(2,0) & " - " & WinPair(2,1) & ", " & WinTres(2,0) & " - " & WinTres(2,1)
	  
	  
	else
	  Grp2 = GrpTemp
	  Grp3 = GrpTemp
	end if
  
    response = 6
    SelLoto
  next
  Grp2 = GrpTempAll
  Grp3 = GrpTempAll
  Grp4 = GrpTempAll
  Grp5 = GrpTempAll
  PickThree = ""
next
end if

'1111 Pattern (AKA Kevin's Numbers) - Revised 11/14/13
'-----------------------
'Grp1 = "11,17,25,32,34,35,40,41,48,49,55,57"
'Grp2 = "7,9,10,13,14,29,36,39,43,46,50,52,54"
'Grp3 = "1,2,6,12,19,21,26,27,44,56"
'Grp4 = "8,15,20,24,31,33,37,38,45,51,53"
'Grp5 = "3,4,5,16,18,22,23,28,30,42,47,58,59"

'for WinCount = 0 to 4 'Remove winning numbers from Quartile
'  Grp1 = Replace("," & Grp1 & ",", "," & WinningNbrs(WinCount) & ",", ",")
'  if left(Grp1, 1) = "," then Grp1 = right(Grp1, len(Grp1)-1)
'  if right(Grp1, 1) = "," then Grp1 = left(Grp1, len(Grp1)-1)
'  
'  Grp2 = Replace("," & Grp2 & ",", "," & WinningNbrs(WinCount) & ",", ",")
'  if left(Grp2, 1) = "," then Grp2 = right(Grp2, len(Grp2)-1)
'  if right(Grp2, 1) = "," then Grp2 = left(Grp2, len(Grp2)-1)
'  
'  Grp3 = Replace("," & Grp3 & ",", "," & WinningNbrs(WinCount) & ",", ",")
'  if left(Grp3, 1) = "," then Grp3 = right(Grp3, len(Grp3)-1)
'  if right(Grp3, 1) = "," then Grp3 = left(Grp3, len(Grp3)-1)
'  
'  Grp4 = Replace("," & Grp4 & ",", "," & WinningNbrs(WinCount) & ",", ",")
'  if left(Grp4, 1) = "," then Grp4 = right(Grp4, len(Grp4)-1)
'  if right(Grp4, 1) = "," then Grp4 = left(Grp4, len(Grp4)-1)
'  
'  Grp5 = Replace("," & Grp5 & ",", "," & WinningNbrs(WinCount) & ",", ",")
'  if left(Grp5, 1) = "," then Grp5 = right(Grp5, len(Grp5)-1)
'  if right(Grp5, 1) = "," then Grp5 = left(Grp5, len(Grp5)-1)
'next

'If commenting this section out, remove comment below random numbers
'-----------------------


'Choose random numbers
Grp1 = GrpTempAll 'removed for Quartile
OutputNote = OutputNote & vbcrlf & "'Random - Differences" & vbcrlf
RndAmt = inputbox("Enter how many RANDOM numbers to pick:", "RANDOM Numbers", "6") 'How many random numbers to choose
if not IsNumeric(RndAmt) then RndAmt = 6
for WinFive = 1 to int(RndAmt)
    response = 6
    SelLoto
next

WriteFile format(date() + SatWed(date()), "YYYY-MM-DD") & "_PB-Numbers-Differences.txt", OutputNote
msgbox "Sets have been created"

Function SelLoto()

'Set Array to all zeros
SetsAr(0) = 0
SetsAr(1) = 0
SetsAr(2) = 0
SetsAr(3) = 0
SetsAr(4) = 0

i = 1
pt= ""

Do Until i > 5

   If i = 1 then ThisGrp = Grp1 else if i = 2 then ThisGrp = Grp2 else if i = 3 then ThisGrp = Grp3 else if i = 4 then ThisGrp = Grp4 else if i = 5 then ThisGrp = Grp5
   
   'Remove Dupes
   for j = 0 to i - 1
     ThisGrp = Replace("," & ThisGrp & ",", "," & SetsAr(j) & ",", ",")
	 if left(ThisGrp, 1) = "," then ThisGrp = right(ThisGrp, len(ThisGrp)-1)
	 if right(ThisGrp, 1) = "," then ThisGrp = left(ThisGrp, len(ThisGrp)-1)
   next
   
   GrpCount = 1 + len(ThisGrp) - len(Replace(ThisGrp, ",", ""))
   
   Randomize   ' Initialize random-number generator.
   MyValue = Int(GrpCount * Rnd) + 1   ' Generate random value between 1 and 6.
   
   'msgbox "Set: " & i & vbcrlf & "Count: " & GrpCount & vbcrlf & "Pick: " & MyValue

   if MyValue > 0 and MyValue =< GrpCount then
	 countComma = 0
	 for j = 0 to len(ThisGrp)
	   if countComma = 0 then
	     countComma = 1
	   else
	     if mid(ThisGrp,j,1) = "," then countComma = countComma + 1
	   end if
	   if countComma = MyValue then
	     posNumber = j + 1
		 j = len(ThisGrp)
	   end if
	 next
	 'msgbox posNumber
	 if posNumber + 2 => len(ThisGrp) then
	   MyValue = mid(ThisGrp,posNumber,len(ThisGrp) - posNumber + 1)
	 else
	   MyValue = mid(ThisGrp,posNumber,InStr(posNumber,ThisGrp,",") - posNumber)
	 end if
	 'msgbox MyValue
	 SetsAr(i-1) = MyValue
	 'if len(pt) > 0 then pt = pt & "-" & MyValue else pt = MyValue
   end if

   'MsgBox MyValue & ", " & pt

   i= i + 1

Loop


'Check numbers - total
if int(SetsAr(0)) + int(SetsAr(1)) + int(SetsAr(2)) + int(SetsAr(3)) + int(SetsAr(4)) < 109 or int(SetsAr(0)) + int(SetsAr(1)) + int(SetsAr(2)) + int(SetsAr(3)) + int(SetsAr(4)) > 187 then
  'msgbox "T<>-Check = " & int(SetsAr(0)) + int(SetsAr(1)) + int(SetsAr(2)) + int(SetsAr(3)) + int(SetsAr(4))
  SelLoto

'Check numbers - Even/Odd
elseif (int(SetsAr(0)) mod 2) + (int(SetsAr(1)) mod 2) + (int(SetsAr(2)) mod 2) + (int(SetsAr(3)) mod 2) + (int(SetsAr(4)) mod 2) = 0 or (int(SetsAr(0)) mod 2) + (int(SetsAr(1)) mod 2) + (int(SetsAr(2)) mod 2) + (int(SetsAr(3)) mod 2) + (int(SetsAr(4)) mod 2) = 5 then
  'msgbox "S=EO-Check = " & int(SetsAr(0)) mod 2 & " " & int(SetsAr(1)) mod 2 & " " & int(SetsAr(2)) mod 2 & " " & int(SetsAr(3)) mod 2 & " " & int(SetsAr(4)) mod 2
  SelLoto
end if

'Sort Array
Sort SetsAr

'Check numbers - 8 thing
checkYN = False
for i = 1 to 5
  Check1 = SetsAr(i-1)
  if not i = 5 then
    Check2 = SetsAr(i)
	if Check2 - Check1 < 5 then checkYN = True
	'msgbox Check2 & " - " & Check1 & " = " & Check2 - Check1 & " " & checkYN
  end if
next
if checkYN = False then
  'msgbox "S>8-Check " & SetsAr(0) & "-" & SetsAr(1) & "-" & SetsAr(2) & "-" & SetsAr(3) & "-" & SetsAr(4)
  SelLoto
else

  'Check numbers - Differences
  Dim CountDiffs
  CountDiffs = 0
  '#1
  if SetsAr(1) - SetsAr(0) >= 1 and SetsAr(1) - SetsAr(0) <= 23 then CountDiffs = CountDiffs + 1
  if SetsAr(2) - SetsAr(0) >= 5 and SetsAr(2) - SetsAr(0) <= 33 then CountDiffs = CountDiffs + 1
  if SetsAr(3) - SetsAr(0) >= 15 and SetsAr(3) - SetsAr(0) <= 47 then CountDiffs = CountDiffs + 1
  if SetsAr(4) - SetsAr(0) >= 21 and SetsAr(4) - SetsAr(0) <= 52 then CountDiffs = CountDiffs + 1
  '#2
  if SetsAr(2) - SetsAr(1) >= 1 and SetsAr(2) - SetsAr(1) <= 22 then CountDiffs = CountDiffs + 1
  if SetsAr(3) - SetsAr(1) >= 7 and SetsAr(3) - SetsAr(1) <= 36 then CountDiffs = CountDiffs + 1
  if SetsAr(4) - SetsAr(1) >= 15 and SetsAr(4) - SetsAr(1) <= 47 then CountDiffs = CountDiffs + 1
  '#3
  if SetsAr(3) - SetsAr(2) >= 1 and SetsAr(3) - SetsAr(2) <= 22 then CountDiffs = CountDiffs + 1
  if SetsAr(4) - SetsAr(2) >= 3 and SetsAr(4) - SetsAr(2) <= 33 then CountDiffs = CountDiffs + 1
  '#4
  if SetsAr(4) - SetsAr(3) >= 1 and SetsAr(4) - SetsAr(3) <= 19 then CountDiffs = CountDiffs + 1
  '#5 - Already taken care of
  
  if CountDiffs < 9 then
    'msgbox "S>Diff-Check " & SetsAr(0) & "-" & SetsAr(1) & "-" & SetsAr(2) & "-" & SetsAr(3) & "-" & SetsAr(4) & vbcrlf & CountDiffs
    SelLoto
  end if
end if
  
if response = 6 then
  DispResult
  response = 0
end if


End Function


'DispResult

Function DispResult()
'Sort Array
'Sort SetsAr

pt = SetsAr(0) & "-" & SetsAr(1) & "-" & SetsAr(2) & "-" & SetsAr(3) & "-" & SetsAr(4)

If filesys.FileExists("PreviousPB.csv") Then
  'PrevWinners = getfile("PreviousPB.csv")
  
  'Check for Previous Winner
  isWin = InStr(1,PrevWinners,Replace(pt, "-", ","))
  if isWin > 0  then
    response = MsgBox(pt & vbcrlf & vbcrlf & "This number was a winner on: " & mid(PrevWinners, isWin - 11, 10) & vbcrlf & vbcrlf & "Try again?", vbyesno)
	if response = 6 then 
	  SelLoto
	end if
  else
  
    'Rank numbers
	RankSetCount = 0
	RankZeros = 0
	RankAll = 0 'The count of all the times a all of the numbers in the sets appeared before
	for i = 0 to ubound(SetsAr)
	  RankComp = PrevWinners
	  RankSets(i) = "" 'Clear Rank Set
	  
	  j = InStr(1,replace(RankComp, "|", ","),"," & SetsAr(i) & ",")
	  do while j > 0
	    FoundRank = mid(RankComp,InStrRev(RankComp,"|",j) + 1, InStr(j,RankComp,"P") - InStrRev(RankComp,"|",j) - 2)
	    'msgbox InStrRev(RankComp,"|",j) & " " & InStr(j,RankComp,"P") & vbcrlf & FoundRank
		
		RankSets(i) = RankSets(i) & "," & FoundRank & ","
		RankComp = replace(RankComp, FoundRank, "0,0,0,0,0")
		
		'Look again
		j = InStr(1,replace(RankComp, "|", ","),"," & SetsAr(i) & ",")
		
		RankAll = RankAll + 1
	  loop
	  if not i = 0 then
	    for j = 0 to i - 1
		  RankTotals(RankSetCount) = (len(RankSets(j)) - len(replace(RankSets(j),"," & SetsAr(i) & "," , ",")))/len(SetsAr(i) & ",")
		  if RankTotals(RankSetCount) = 0 then RankZeros = RankZeros + 1
		  RankSetCount = RankSetCount + 1
		next
	  end if
    next
	
	'Check if there were more than 1 bad ranking number
	if RankZeros > 1 then
	  response = 6
	  'msgbox "Bad number found. Retrying..."
	else
      response = 0
	  OutputNote = OutputNote & pt & " (" & _
	    100*round((RankTotals(0) + RankTotals(1) + RankTotals(2) + RankTotals(3) + RankTotals(4) + RankTotals(5) + RankTotals(6) + RankTotals(7) + RankTotals(8) + RankTotals(9))/RankAll,2) & _
		")" & vbcrlf
	  
	  'response = MsgBox(vbtab & vbtab & "Your number is:" & vbcrlf & vbcrlf & _ 
	  ' vbtab & vbtab & pt & vbcrlf & vbcrlf & _
	  ' vbtab & SetsAr(0) & vbtab & SetsAr(1) & vbtab & SetsAr(2) & vbtab & SetsAr(3) & vbtab & SetsAr(4) & vbcrlf & vbcrlf & _
	  ' SetsAr(0) & vbtab & " " & vbtab & RankTotals(0) & vbtab & RankTotals(1) & vbtab & RankTotals(3) & vbtab & RankTotals(6) & vbcrlf & _
	  ' SetsAr(1) & vbtab & " " & vbtab & " " & vbtab & RankTotals(2) & vbtab & RankTotals(4) & vbtab & RankTotals(7) & vbcrlf & _
	  ' SetsAr(2) & vbtab & " " & vbtab & " " & vbtab & " " & vbtab & RankTotals(5) & vbtab & RankTotals(8) & vbcrlf & _
	  ' SetsAr(3) & vbtab & " " & vbtab & " " & vbtab & " " & vbtab & " " & vbtab & RankTotals(9) &  vbcrlf & _
	  ' SetsAr(4) & vbtab & " " & vbtab & " " & vbtab & " " & vbtab & " " & vbtab & " " & vbcrlf & vbcrlf & _
	  ' "Score is: " & RankTotals(0) + RankTotals(1) + RankTotals(2) + RankTotals(3) + RankTotals(4) + RankTotals(5) + RankTotals(6) + RankTotals(7) + RankTotals(8) + RankTotals(9) & vbcrlf & vbcrlf & _
	  ' "Try again?", vbyesno)
	 end if
	 if response = 6 then 
	    SelLoto
	 end if
  end if
else
  response = MsgBox(pt & vbcrlf & vbcrlf & "Try again?", vbyesno)
  if response = 6 then 
	  SelLoto
	end if
end if

'Clipboard copy
'Set objIE = CreateObject("InternetExplorer.Application")
'objIE.Navigate("about:blank")
'objIE.document.parentwindow.clipboardData.SetData "text", pt
'objIE.Quit

End Function



'Sort Array by Number
Sub Sort( ByRef myArray )
    Dim i, j, strHolder

    For i = ( UBound( myArray ) - 1 ) to 0 Step -1
        For j= 0 to i
            If int( myArray( j ) ) > int( myArray( j + 1 ) ) Then
                strHolder        = myArray( j + 1 )
                myArray( j + 1 ) = myArray( j )
                myArray( j )     = strHolder
            End If
        Next
    Next 
End Sub

'Sort Array by Number2
Sub SortB( ByRef myArray )
    Dim i, j, strHolder

    For i = ( UBound( myArray ) - 1 ) to 0 Step -1
        For j= 0 to i
            If int( myArray(j,1) ) < int( myArray(j + 1,1) ) Then
                strHolder        = myArray(j + 1,1)
                myArray(j + 1,1) = myArray(j,1)
                myArray(j,1)     = strHolder
				
				strHolder        = myArray(j + 1,0)
                myArray(j + 1,0) = myArray(j,0)
                myArray(j,0)     = strHolder
            End If
        Next
    Next 
End Sub

'Read text file
function GetFile(FileName)
  If FileName<>"" Then
    Dim FS, FileStream
    Set FS = CreateObject("Scripting.FileSystemObject")
      on error resume Next
      Set FileStream = FS.OpenTextFile(FileName)
      GetFile = FileStream.ReadAll
  End If
End Function

'Write string As a text file.
function WriteFile(FileName, Contents)
  Dim OutStream, FS

  on error resume Next
  Set FS = CreateObject("Scripting.FileSystemObject")
    Set OutStream = FS.OpenTextFile(FileName, 2, True)
    OutStream.Write Contents
End Function

'Format date/time
Function Format(vExpression, sFormat) 
  set fmt = CreateObject("MSSTDFMT.StdDataFormat") 
  fmt.Format = sFormat 
 
  set rsf = CreateObject("ADODB.Recordset") 
  rsf.Fields.Append "fldExpression", 12 ' adVariant 
 
  rsf.Open 
  rsf.AddNew 
 
  set rsf("fldExpression").DataFormat = fmt 
  rsf("fldExpression").Value = vExpression 
 
  Format = rsf("fldExpression").Value 
 
  rsf.close: Set rsf = Nothing: Set fmt = Nothing 
 
End Function

'Generate next Wednesday/Saturday
Function SatWed(mydate)
  if format(mydate, "Ddd") = "Sat" or format(mydate, "Ddd") = "Wed" then
    SatWed = 0
  elseif format(mydate, "Ddd") = "Sun" then
    SatWed = 3
  elseif format(mydate, "Ddd") = "Mon" then
    SatWed = 2
  elseif format(mydate, "Ddd") = "Tue" then
    SatWed = 1
  elseif format(mydate, "Ddd") = "Thu" then
    SatWed = 2
  elseif format(mydate, "Ddd") = "Fri" then
    SatWed = 1
  end if
end function