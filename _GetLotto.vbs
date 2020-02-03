Dim i, j, PrevWinners, LocalWinners, LastWin, response
Dim PBnums(4) 'The last set of winning numbers
Dim PowerBall 'Powerball Number
Dim NumbersPlayed 'Numbers played for last drawing
Dim NPset, CurNum, Count4Set, TotalHits, Hit2, Hit3, HitPB, TempWin, TotalWin 'Counts for hits and winnings
Dim WinHTML, BoldHTML 'HTML version of winnings for email
dim xmlhttp : set xmlhttp = createobject("msxml2.xmlhttp.3.0")

'Set integers to zero
TotalHits = 0
Hit2 = 0
Hit3 = 0
HitPB = 0
TotalWin = 0

'Get Numbers Online
xmlhttp.open "get", "http://www.powerball.com/powerball/winnums-text.txt", false
xmlhttp.send
PrevWinners = xmlhttp.responseText
set xmlhttp = nothing

'Get local numbers
LocalWinners = getfile("C:\LargeFolder\Powerball\PreviousPB.csv")

'Get Played numbers
NumbersPlayed = getfile("C:\LargeFolder\Powerball\NumbersPlayed.txt")

If len(PrevWinners) > 41 then
  LastWin = mid(PrevWinners,41,35)

  'WriteFile "C:\LargeFolder\Powerball\Lotto.txt", LastWin

  'if 1 = 1 then 'Uncomment for testing and comment next line
  if left(LastWin,10) = format(date() - 1, "MM/DD/YYYY") then
    PBnums(0) = int(mid(LastWin,13,2))
    PBnums(1) = int(mid(LastWin,17,2))
    PBnums(2) = int(mid(LastWin,21,2))
    PBnums(3) = int(mid(LastWin,25,2))
    PBnums(4) = int(mid(LastWin,29,2))
    PowerBall = int(mid(LastWin,33,2))
  
    'Sort and update LastWin
    Sort PBnums 'Sort sets
    LastWin = left(LastWin,10) & "|" & PBnums(0) & "," & PBnums(1) & "," & PBnums(2) & "," & PBnums(3) & "," & PBnums(4) & ",P" & PowerBall
  
    'msgbox LastWin
    WriteFile "C:\LargeFolder\Powerball\PreviousPB.csv", LastWin & vbcrlf & LocalWinners
	
	'Testing Fake Numbers!!
    'PBnums(0) = 2
    'PBnums(1) = 11
    'PBnums(2) = 35
    'PBnums(3) = 52
    'PBnums(4) = 54
    'PowerBall = 13
  
    'Check played numbers and send results e-mail
	If len(NumbersPlayed) > 20 then
	  Do while len(NumbersPlayed) > 20
        Count4Set = 0
	    NPset = left(NumbersPlayed, 20)
	
	    i = -2
	    Do while i < 13
	      i = i + 3 '1,4,7,10,13 (2 spaces)
		  CurNum = mid(NPset,i,2)
	      For j = 0 to ubound(PBnums)
	        if int(CurNum) = PBnums(j) then
		      Count4Set = Count4Set + 1
			  BoldHTML = 1
		    end if
	      Next
		  if BoldHTML = 1 then
		    WinHTML = WinHTML & "<font color=0000FF><b>" & CurNum & "</b></font> &nbsp;"
		    BoldHTML = 0
		  else
		    WinHTML = WinHTML & CurNum & " &nbsp;"
		  end if
	    loop
	    if int(right(NPset,2)) = PowerBall then
	      HitPB = HitPB + 1
		  WinHTML = WinHTML & "<font color=0000FF><b>PB-" & right(NPset,2) & "</b></font>"
	      TempWin = GetWinnings(Count4Set,1)
	    else
	      WinHTML = WinHTML & "PB-" & right(NPset,2)
	      TempWin = GetWinnings(Count4Set,0)
	    End if
	    if TempWin > 0 then
	      WinHTML = WinHTML & "<font color=0000FF><b> - $" & TempWin & " WINNER</b></font><br>" & vbcrlf
	    else
	      WinHTML = WinHTML & "<br>" & vbcrlf
	    End if
	    TotalWin = TotalWin + TempWin
	
	    if Count4Set = 2 then Hit2 = Hit2 + 1
	    if Count4Set = 3 then Hit3 = Hit3 + 1
	    TotalHits = TotalHits + Count4Set
	
	    'msgbox NPset & vbcrlf & "Hits: " & Count4Set & " TotalPB: " & HitPB
	    NumbersPlayed = right(NumbersPlayed,len(NumbersPlayed)-22)
      loop
	
	  'Add fluff text and stats
	  WinHTML = "<html><body>Powerball results for drawing on " & format(date() - 1, "MM/DD/YYYY") & ":<br><br>" & vbcrlf & vbcrlf &_
	    "The winning numbers were <font color=0000FF><b>" & GetTD(PBnums(0)) & " " & GetTD(PBnums(1)) & " " & GetTD(PBnums(2)) & " " & GetTD(PBnums(3)) & " " & GetTD(PBnums(4)) & " PB-" & GetTD(PowerBall) & "</b></font>, " &_
	    "we won $" & TotalWin & "<br><br>" & vbcrlf & vbcrlf &_
	    "Two in set: " & Hit2 & vbcrlf & "<br>Three in set: " & Hit3 & vbcrlf & "<br>Total PB: " & HitPB & vbcrlf & "<br>Total Hits: " & TotalHits & "<br><br>" & vbcrlf & vbcrlf & WinHTML
	  WinHTML = WinHTML & "</body></html>"
	
	  'Test for HTML email
	  'WriteFile "C:\LargeFolder\Powerball\Email.htm", WinHTML
  
      'msgbox "Winnings: $" & TotalWin & vbcrlf & "Pairings: " & Hit2 & " Three: " & Hit3 & " TotalHits: " & TotalHits & " TotalPB: " & HitPB
      if TotalWin > 4 then 'Winning threshold for text notification
        'Text alert for just me:
		SendMail "1234567890@vtext.com", "Large PB result", "Winnings: $" & TotalWin & vbcrlf & "Pairings: " & Hit2 & " Three: " & Hit3 & vbcrlf & "TotalHits: " & TotalHits & " TotalPB: " & HitPB, 0
      end if
	  'Use this for group:
	  SendMail "powerball@someplace.com", "Powerball results for " & format(date() - 1, "MM/DD/YYYY"), WinHTML, 1
    end if
  else
   'msgbox "There was no drawing today"
  End if
End if


'Count dollars won
Function GetWinnings(NonPB, PByn)
  Dim WinAmt
  WinAmt = 0
  
  if PByn = 0 then 'No Powerball
    if NonPB = 3 then
	  WinAmt = 7
	elseif NonPB = 4 then
	  WinAmt = 100
	elseif NonPB = 5 then
	  WinAmt = 1000000
	else
	  WinAmt = 0
	end if
  else 'Yes Powerball
    if NonPB = 2 then
	  WinAmt = 7
	elseif NonPB = 3 then
	  WinAmt = 100
	elseif NonPB = 4 then
	  WinAmt = 50000
	elseif NonPB = 5 then
	  WinAmt = 10000000000
	else
	  WinAmt = 4
	end if
  End if
  
  GetWinnings = WinAmt
end function

'Get two digit number again
Function GetTD(PBnum)
  if PBnum < 10 then
    GetTD = "0" & PBnum
  else
    GetTD = PBnum
  end if
end function

Function SendMail(TextRcv,TextSubject,TextBody,HTMLyn)
  Const cdoSendUsingPickup = 1 'Send message using the local SMTP service pickup directory. 
  Const cdoSendUsingPort = 2 'Send the message using the network (SMTP over the network). 

  Const cdoAnonymous = 0 'Do not authenticate
  Const cdoBasic = 1 'basic (clear-text) authentication
  Const cdoNTLM = 2 'NTLM

  Set objMessage = CreateObject("CDO.Message") 
  objMessage.Subject = TextSubject 
  objMessage.From = """PB Mailer"" <powerball@someplace.com>" 
  objMessage.To = TextRcv
  if HTMLyn = 1 then
    objMessage.HTMLBody = TextBody
  else
    objMessage.TextBody = TextBody
  end if

  '==This section provides the configuration information for the remote SMTP server.

  objMessage.Configuration.Fields.Item _
  ("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2 

  'Name or IP of Remote SMTP Server
  objMessage.Configuration.Fields.Item _
  ("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "mail.someplace.com"

  'Type of authentication, NONE, Basic (Base64 encoded), NTLM
  objMessage.Configuration.Fields.Item _
  ("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = cdoBasic

  'Your UserID on the SMTP server
  objMessage.Configuration.Fields.Item _
  ("http://schemas.microsoft.com/cdo/configuration/sendusername") = "username"

  'Your password on the SMTP server
  objMessage.Configuration.Fields.Item _
  ("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "password"

  'Server port (typically 25)
  objMessage.Configuration.Fields.Item _
  ("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25

  'Use SSL for the connection (False or True)
  objMessage.Configuration.Fields.Item _
  ("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = False

  'Connection Timeout in seconds (the maximum time CDO will try to establish a connection to the SMTP server)
  objMessage.Configuration.Fields.Item _
  ("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60

  objMessage.Configuration.Fields.Update

  '==End remote SMTP server configuration section==

  objMessage.Send 
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