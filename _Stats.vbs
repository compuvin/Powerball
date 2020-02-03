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
dim TestSets(4), Pattern, Ranking, Quartile, Differences 'Strings for each test/pattern
dim filesys
set filesys=CreateObject("Scripting.FileSystemObject")

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

for WinCount = 0 to 4
  WinningNbrs(WinCount) = int(left(LastLotto,instr(1,LastLotto,",")-1))
  LastLotto = replace("," & LastLotto,"," & WinningNbrs(WinCount) & ",", "")
  'msgbox WinningNbrs(WinCount)
next

'Pattern
For i = 0 to 4
  TestSets(i) = 0
next
Grp1 = "11,17,25,32,34,35,40,41,48,49,55,57"
Grp2 = "7,9,10,13,14,29,36,39,43,46,50,52,54"
Grp3 = "1,2,6,12,19,21,26,27,44,56"
Grp4 = "8,15,20,24,31,33,37,38,45,51,53"
Grp5 = "3,4,5,16,18,22,23,28,30,42,47,58,59"
for WinCount = 0 to 4
  TestSets(0) = TestSets(0) + ((len("," & Grp1 & ",") - len(replace("," & Grp1 & ",","," & WinningNbrs(WinCount) & "," , ",")))/len(WinningNbrs(WinCount) & ","))
  TestSets(1) = TestSets(1) + ((len("," & Grp2 & ",") - len(replace("," & Grp2 & ",","," & WinningNbrs(WinCount) & "," , ",")))/len(WinningNbrs(WinCount) & ","))
  TestSets(2) = TestSets(2) + ((len("," & Grp3 & ",") - len(replace("," & Grp3 & ",","," & WinningNbrs(WinCount) & "," , ",")))/len(WinningNbrs(WinCount) & ","))
  TestSets(3) = TestSets(3) + ((len("," & Grp4 & ",") - len(replace("," & Grp4 & ",","," & WinningNbrs(WinCount) & "," , ",")))/len(WinningNbrs(WinCount) & ","))
  TestSets(4) = TestSets(4) + ((len("," & Grp5 & ",") - len(replace("," & Grp5 & ",","," & WinningNbrs(WinCount) & "," , ",")))/len(WinningNbrs(WinCount) & ","))
next
Pattern = 0
for WinCount = 0 to 4
  if not TestSets(WinCount) = 0 then Pattern = Pattern + 1
next
OutputNote = "Pattern: " & TestSets(0) & TestSets(1) & TestSets(2) & TestSets(3) & TestSets(4) & " = " & (Pattern/5)*100 & "%" & vbcrlf

'Ranking
Grp1 = "1,3,4,6,8,9,10,11,12,13,16,17,18,19,27,28,37,38,39,40,42,44,46,47,49,50,51,53,54,56,59" 'Yellow numbers
Grp2 = Grp1 'Yellow numbers
Grp3 = Grp1 'Yellow numbers
Grp4 = Grp1 'Yellow numbers
Grp5 = "2,5,7,14,15,20,21,22,23,24,25,26,29,30,31,32,33,34,35,36,41,43,45,48,52,55,57,58" 'All numbers (was Grp1 & "," & ...)
Ranking = 0
for WinCount = 0 to 4
  Ranking = Ranking + ((len("," & Grp5 & ",") - len(replace("," & Grp5 & ",","," & WinningNbrs(WinCount) & "," , ",")))/len(WinningNbrs(WinCount) & ","))
next
if Ranking > 1 then
  OutputNote = OutputNote &  "Ranking: " & ((6 - Ranking)/5)*100 & "%" & vbcrlf
else
  OutputNote = OutputNote & "Ranking: 100%" & vbcrlf
end if

'Quartile
For i = 0 to 4
  TestSets(i) = 0
next
Grp1 = "1,2,3,4,5,6,7,8,9,10,11,12"
Grp2 = Grp1 & ",13,14,16,17,18,19,20,21,22,23,24"
Grp3 = Grp2 & ",26,27,28,30,32,33,34,35"
Grp4 = Grp3 & ",36,37,38,39,40,41,42,43,44,45,46,47,48,49"
Grp5 = Grp4 & ",50,51,52,53,54,55,56,57,58,59" & ",15,25,29,31" 'Less of a chance numbers
for WinCount = 0 to 4
  TestSets(0) = TestSets(0) + ((len("," & Grp1 & ",") - len(replace("," & Grp1 & ",","," & WinningNbrs(WinCount) & "," , ",")))/len(WinningNbrs(WinCount) & ","))
  TestSets(1) = TestSets(1) + ((len("," & Grp2 & ",") - len(replace("," & Grp2 & ",","," & WinningNbrs(WinCount) & "," , ",")))/len(WinningNbrs(WinCount) & ","))
  TestSets(2) = TestSets(2) + ((len("," & Grp3 & ",") - len(replace("," & Grp3 & ",","," & WinningNbrs(WinCount) & "," , ",")))/len(WinningNbrs(WinCount) & ","))
  TestSets(3) = TestSets(3) + ((len("," & Grp4 & ",") - len(replace("," & Grp4 & ",","," & WinningNbrs(WinCount) & "," , ",")))/len(WinningNbrs(WinCount) & ","))
  TestSets(4) = TestSets(4) + ((len("," & Grp5 & ",") - len(replace("," & Grp5 & ",","," & WinningNbrs(WinCount) & "," , ",")))/len(WinningNbrs(WinCount) & ","))
next
Quartile = 0
for WinCount = 0 to 4
  if not TestSets(WinCount) = 0 then Quartile = Quartile + 1
next
OutputNote = OutputNote &  "Quartile: " & (Quartile/5)*100 & "%" & vbcrlf

'Differences
'#1
if WinningNbrs(1) - WinningNbrs(0) >= 1 and WinningNbrs(1) - WinningNbrs(0) <= 23 then Differences = Differences + 1
if WinningNbrs(2) - WinningNbrs(0) >= 5 and WinningNbrs(2) - WinningNbrs(0) <= 33 then Differences = Differences + 1
if WinningNbrs(3) - WinningNbrs(0) >= 15 and WinningNbrs(3) - WinningNbrs(0) <= 47 then Differences = Differences + 1
if WinningNbrs(4) - WinningNbrs(0) >= 21 and WinningNbrs(4) - WinningNbrs(0) <= 52 then Differences = Differences + 1
'#2
if WinningNbrs(2) - WinningNbrs(1) >= 1 and WinningNbrs(2) - WinningNbrs(1) <= 22 then Differences = Differences + 1
if WinningNbrs(3) - WinningNbrs(1) >= 7 and WinningNbrs(3) - WinningNbrs(1) <= 36 then Differences = Differences + 1
if WinningNbrs(4) - WinningNbrs(1) >= 15 and WinningNbrs(4) - WinningNbrs(1) <= 47 then Differences = Differences + 1
'#3
if WinningNbrs(3) - WinningNbrs(2) >= 1 and WinningNbrs(3) - WinningNbrs(2) <= 22 then Differences = Differences + 1
if WinningNbrs(4) - WinningNbrs(2) >= 3 and WinningNbrs(4) - WinningNbrs(2) <= 33 then Differences = Differences + 1
'#4
if WinningNbrs(4) - WinningNbrs(3) >= 1 and WinningNbrs(4) - WinningNbrs(3) <= 19 then Differences = Differences + 1
'#5 - Already taken care of
  
if Differences < 9 then
  OutputNote = OutputNote &  "Differences: " & Differences & " = 0%" & vbcrlf
else
  OutputNote = OutputNote &  "Differences: " & Differences & " = 100%" & vbcrlf
end if

'Rank numbers
RankSetCount = 0
RankZeros = 0
RankAll = 0 'The count of all the times a all of the numbers in the sets appeared before
for i = 0 to ubound(WinningNbrs)
  RankComp = replace(PrevWinners, mid(PrevWinners,instr(1,PrevWinners,"|")+1,instr(1,PrevWinners,",P")-instr(1,PrevWinners,"|")), "0,0,0,0,0")
  RankSets(i) = "" 'Clear Rank Set
 
  j = InStr(1,replace(RankComp, "|", ","),"," & WinningNbrs(i) & ",")
  do while j > 0
    FoundRank = mid(RankComp,InStrRev(RankComp,"|",j) + 1, InStr(j,RankComp,"P") - InStrRev(RankComp,"|",j) - 2)
    'msgbox InStrRev(RankComp,"|",j) & " " & InStr(j,RankComp,"P") & vbcrlf & FoundRank
		
	RankSets(i) = RankSets(i) & "," & FoundRank & ","
	RankComp = replace(RankComp, FoundRank, "0,0,0,0,0")
		
	'Look again
	j = InStr(1,replace(RankComp, "|", ","),"," & WinningNbrs(i) & ",")
		
	RankAll = RankAll + 1
  loop
  if not i = 0 then
    for j = 0 to i - 1
	  RankTotals(RankSetCount) = (len(RankSets(j)) - len(replace(RankSets(j),"," & WinningNbrs(i) & "," , ",")))/len(WinningNbrs(i) & ",")
	  if RankTotals(RankSetCount) = 0 then RankZeros = RankZeros + 1
	  RankSetCount = RankSetCount + 1
	next
  end if
next
	
OutputNote = OutputNote & "Score: " & _
100*round((RankTotals(0) + RankTotals(1) + RankTotals(2) + RankTotals(3) + RankTotals(4) + RankTotals(5) + RankTotals(6) + RankTotals(7) + RankTotals(8) + RankTotals(9))/RankAll,2) &_
" (" & RankZeros & " zeros)" & vbcrlf


msgbox "Stats for last drawing:" & vbcrlf & vbcrlf & OutputNote

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
  Dim nExpression
  nExpression = sFormat
  
  if isnull(vExpression) = False then
    if instr(1,sFormat,"Y") > 0 or instr(1,sFormat,"M") > 0 or instr(1,sFormat,"D") > 0 or instr(1,sFormat,"H") > 0 or instr(1,sFormat,"S") > 0 then 'Time/Date Format
      vExpression = cdate(vExpression)
	  if instr(1,sFormat,"AM/PM") > 0 and int(hour(vExpression)) > 12 then
	    nExpression = replace(nExpression,"HH",right("00" & hour(vExpression)-12,2)) '2 character hour
	    nExpression = replace(nExpression,"H",hour(vExpression)-12) '1 character hour
		nExpression = replace(nExpression,"AM/PM","PM") 'If if its afternoon, its PM
	  else
	    nExpression = replace(nExpression,"HH",right("00" & hour(vExpression),2)) '2 character hour
	    nExpression = replace(nExpression,"H",hour(vExpression)) '1 character hour
		if int(hour(vExpression)) = 12 then nExpression = replace(nExpression,"AM/PM","PM") '12 noon is PM while anything else in this section is AM (fixed 04/19/2019 thanks to our HR Dept.)
		nExpression = replace(nExpression,"AM/PM","AM") 'If its not PM, its AM
	  end if
	  nExpression = replace(nExpression,":MM",":" & right("00" & minute(vExpression),2)) '2 character minute
	  nExpression = replace(nExpression,"SS",right("00" & second(vExpression),2)) '2 character second
	  nExpression = replace(nExpression,"YYYY",year(vExpression)) '4 character year
	  nExpression = replace(nExpression,"YY",right(year(vExpression),2)) '2 character year
	  nExpression = replace(nExpression,"DD",right("00" & day(vExpression),2)) '2 character day
	  nExpression = replace(nExpression,"D",day(vExpression)) '(N)N format day
	  nExpression = replace(nExpression,"MMM",left(MonthName(month(vExpression)),3)) '3 character month name
	  if instr(1,sFormat,"MM") > 0 then
	    nExpression = replace(nExpression,"MM",right("00" & month(vExpression),2)) '2 character month
	  else
	    nExpression = replace(nExpression,"M",month(vExpression)) '(N)N format month
	  end if
    elseif instr(1,sFormat,"N") > 0 then 'Number format
	  nExpression = vExpression
	  if instr(1,sFormat,".") > 0 then 'Decimal format
	    if instr(1,nExpression,".") > 0 then 'Both have decimals
		  do while instr(1,sFormat,".") > instr(1,nExpression,".")
		    nExpression = "0" & nExpression
		  loop
		  if len(nExpression)-instr(1,nExpression,".") >= len(sFormat)-instr(1,sFormat,".") then
		    nExpression = left(nExpression,instr(1,nExpression,".")+len(sFormat)-instr(1,sFormat,"."))
	      else
		    do while len(nExpression)-instr(1,nExpression,".") < len(sFormat)-instr(1,sFormat,".")
			  nExpression = nExpression & "0"
			loop
	      end if
		else
		  nExpression = nExpression & "."
		  do while len(nExpression) < len(sFormat)
			nExpression = nExpression & "0"
		  loop
	    end if
	  else
		do while len(nExpression) < sFormat
		  nExpression = "0" and nExpression
		loop
	  end if
	else
      response.write "Formating issue on page. Unrecognized format: " & sFormat
	end if
	
	Format = nExpression
  end if
End Function