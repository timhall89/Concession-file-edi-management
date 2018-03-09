'*****************************************************************************************************
' Script Notes:
' This is the main script for executing the tasks for processing, RPO requests, Aperak responses and ASN's.
' I have contained all 3 main area of operation in one script for simplicity.
' This script can be broken down in to the Following main Sections, in the order which they are found:
'	1) Declaration of script level constants and variables.
'	2) Assigning initial values to some script level Variables.
'	3) The section of code that is run to determine what process is to be run.
'	4) The 3 sub procedures that correspond to each of the 3 main operation of this application as
'	   detailed above.
'	5) All the custom function written to be used in the initial script code and 3 main sub procedures.
' Access import Note: whenever there is refernece to an access import of any kind, it is likely that
' this is not stricly an access import as any "importing" of data for access to use is done with linked
' tables to .txt or .csv files, but the result is the same as importing, so for the sake of clarity,
' there may be some reference to imports in to access in the notes in this script that actualy refer to
' linked tables in access.
'*****************************************************************************************************

'=====================================================================================================
'=====================================================================================================

'*****************************************************************************************************
' Declaration of script level constants.
'*****************************************************************************************************
Option Explicit

' Const of email recipients for admin email.
	Const strAdminEmails = ""

' The email address that all email are sent from.
	Const strEmailFrom = ""

' Path to the folder where the new ASN files are put.
	Const strASNFolder = ""

' Path to the folder where the APERAK files put by david jones.
	Const strAperakPath = ""

' Path to the folder where the RPO files are Archived.
	Const strRpoArchivePth = ""

' Name of the text file containing a list of the invoice numbers to create RPO's for.
	Const strInvNumTxtFname = ""

' Name of the text file containg the invoice information formated for Access to use.
' Note: This text file is a linked table in the Access DataBase.
	Const strImpDataFNm = ""

' File name of the text file that the raw Nova SpoolX invoice information extract is saved to.
	Const strNovaExtRawFNm = ""

' File name of the RPO process error report (Text File).
	Const strRpoErrRepFNm = ""

' File name of the Text file used as the main body of the the RPO creation email response.
	Const strRpoEResFNm = ""

' Number of Available RPO Numbers at which a reminder email is to be sent.
	Const intRPORemaining = 100

' Consts used for the Wait for process end sub.
	Const intIntervalMS = 2000
	Const insIntervalMaxCount = 180

'=====================================================================================================
'=====================================================================================================

'*****************************************************************************************************
' Declaration of script level variables.
'*****************************************************************************************************
	
	Dim mstrTo, mstrCC, mstrBCC, sPath, sName, strAccFullName, strEmailsFullName, dtDTStamp
	Dim argSelect, strEResponseBody, strEmailSubject
	Dim ff, blnDontOutputEnd
	Dim objFSO, objShell, bltTestMode

'*****************************************************************************************************
' Set the values of some script level variables.
'*****************************************************************************************************

' Set the Shell and File system Objects that are used throughout the entire stript.
' Note: There is no need to have more that one instance of these 2 objects at any time, so these
' variable are just used repeatedly for simplicity. 
	Set objShell = CreateObject("WScript.Shell")
	Set objFSO = CreateObject("Scripting.FileSystemObject")

' Get the full path and name of this VB Script.
	sName = WScript.ScriptFullName

' Get just the path of this script including a trailing "\".
	sPath = left(sName, instrRev(sName, "\"))

' Set the Current working directory to where this script is.
	objShell.CurrentDirectory = sPath

' Get the full Name (incl Path) of the Access datata based used by the AutoDJ application.
' This should always be located in the same place as this script.
	strAccFullName = sPath & ""

' Get the full name (incl Path) of the text file containing the email rcepienst for any response emails.
' This should always be located in the same place as this script.
	strEmailsFullName = sPath & ""

' Read in to a variable the Current date and time stamp used throughout this script for prefixing to
' archive file names.
	dtDTStamp = GetDTStamp(Now, True)


'=====================================================================================================
'=====================================================================================================

'*****************************************************************************************************
' Inital code to determine what sub routine to run.
' This is done through the "subCall" script input argument, if this is not supplied (i.e. the vbscript
' is run from windows explorer) then the user will be prompted to input an argument.
'*****************************************************************************************************
	'OutputToLogFile "---------------------------------------"
	'OutputToLogFile "Script Called."
	
' Read the script argument "subCall" In to a varibale.
	argSelect = WScript.Arguments.Named.Item("subCall")
	
' If no argument was suplied it will return an empty string.
	If argSelect = "" Then 
		
	' A pop up is first displayed for 5 seconds and requires user respons to continue to the Input box.
	' The reason for this what is an automated task somehow called the vbscript without passing an Argument
	' the script would just terminate rather than waiting indefinitly for a user in input something.
		If Not objShell.Popup("No input Argument", 5, "AutoDJ") = -1 Then
			argSelect = InputBox("Enter Input Argument", "AutoDJ")
			
			'OutputToLogFile "Argument Input box Displayed"
			
		End If
		
	End If
	
' Wait until no other AutoDJ processess are running.
	WaitForProcessEnd True
	
' A particular case is selected based on the subCall input Argument or the user inputed argument.
	Select Case argSelect
	
	' Case to process RPO Requests.
		Case "RPO"
			RpoProcessing
			'OutputToLogFile "RpoProcessing sub Finished."
			
	' Case to Process ASN's.
		Case "ASN"
			ASNProcessing
			'OutputToLogFile "ASNProcessing sub Finished."
			
	' Case to process Aperaks.
		Case "APK"
			AperakProcessing
			'OutputToLogFile "AperakProcessing sub Finished."
			
	' Case to export EAN list csv file
		Case "EAN"
			ExportEANs
			'OutputToLogFile "ExportEANs sub Finished."
			
	' If the input argument is "" after the manual input process then nothing is to be run and the
	' script is terminated.
		Case ""
			
			'OutputToLogFile "No subCall Argument Supplied or entered."
			blnDontOutputEnd = True
			'WScript.Quit
	
	' Case used to delete old entries of log File.
		Case "LOG"
			
		' Call sub, input is number of days until today to keep
		' ***Number of days to keep can be changed here****
			DeleteLogFileEntries 28

	' Case for calling the sub to check if there are any anomalies.
		Case "ANOM"
			CheckForAnom

	' Case used as a dummy test Case.
		Case "fff"
			'DoCopyMoveFiles strASNFolder & "\ASNData380708.csv", strASNFolder & "\ProcessingFailure", "MOVE"
				' Run an FTP command to send any RPO's file created by access to the David Jones server.
			ExecFTP sPath, "f", "", "", "", "RPO-*.csv", "mput", False
	
	' Used for testeing running on IFT01
		Case "IFTP"
			testIFT01
			
		Case "RRR"
			ExecFTP sPath, "", "", "", "", "RPO-*.csv", "mput", False
	
	' If anything else then a pop message warns that the input is not recognised.
		Case Else
		
			'OutputToLogFile "Argument supplied is not recognised."
			
			objShell.Popup "Argument not recognised", 5, "AutoDJ"
			blnDontOutputEnd = True
			
	End Select
	
	If Not blnDontOutputEnd Then OutputToLogFile "End." & vbCrLf & _
		"------------------------------------------------------------------------------"
	
' Delete the process running indication file.
	DeleteProcessRunning
	
	
' End of script code, below are subs and functions.

'=====================================================================================================
'=====================================================================================================


'*****************************************************************************************************
' Porpose:	Keeps checking if a proccess is running and waiting each time for a set interval.
' Note:		The constants used for interval and max number of loops are set in the script level const
'			declaration section.
'*****************************************************************************************************
Sub WaitForProcessEnd(blnSetProccess)
	
	On Error Resume next
	
	Dim intCounter
	Do While IsProcessRunning And intCounter < insIntervalMaxCount
	
		WScript.Sleep intIntervalMS
		intCounter = intCounter + 1
		If intCounter >= insIntervalMaxCount Then _
		OutputToLogFile "Wait For Process End timed out: " & _
			intIntervalMS * intCounter / 1000 & " Seconds"	
	Loop
		
	If blnSetProccess Then SetProcessRunning
	
End Sub

'*****************************************************************************************************
' Porpose:		Check for any RPO/ASN anomalies with regards to time taken to accept and send an email
'				if any are found.
'*****************************************************************************************************
Sub CheckForAnom()

	On Error Resume Next
	
' Delete info file is exists.
	objFSO.DeleteFile sPath & "AnomFound.info"

' Run the access macro to check for anomalies.
	RunAccessMacro strAccFullName, "CheckForAnoms", 0, True

' If the info file exists indicating that some anomalies were found then send as email.
	If objFSO.FileExists(sPath & "AnomFound.info") Then
		
		Dim acnt
		With objFSO.OpenTextFile(sPath & "AnomFound.info", 1, False)
			acnt = CInt(.ReadLine)
			.Close
		End With
	' Send the email using the custom function for sending an email using the blat application.
	' The email is to be sent to Admin only as a prompt that some anomalies hav been found.
		SendBlatEmail strEmailFrom, strAdminEmails, "", "", _
				"AutoDJ:RPO/ASN Anomalies", acnt & " DJ Deliveies have taken longer than the specified " & _
					"limit to have either the RPO or ASN accepted." & vbCrLf & vbCrLf & _
					"Please check these by going to the David Jones RPO Menu in the REISS EDI Database " & _
					"and Clicking the View Anomaloies Button", False, ""
					
	End If

' Delete the info File
	objFSO.DeleteFile sPath & "AnomFound.info"

' Output message to log file.
	OutputToLogFile "Anomalies Checked."
	
End Sub

'*****************************************************************************************************
' Porpose:		Delete any line in the log file that are earlier than the supplied date.
' Arguments:	strDate: the date, as a string, for which any entries before this are to be deleted.
'				Should be in dd/mm/yyyy format.
'*****************************************************************************************************
Sub DeleteLogFileEntries(intKeepDays)

	On Error Resume Next

' Declare variables.
	Dim objLogFile, objTempLogFile, strLine, blnCopy
	
' Create a temp copy of the log file.
	objFSO.CopyFile sPath & "AutoDJ.log", sPath & "AutoDJTemp.log"

' Open the copy of the log file and set to object variable.
	Set objTempLogFile = objFSO.OpenTextFile(sPath & "AutoDJTemp.log", 1, True)

' Open the original log file for output (will overwrite any exiting data).
	Set objLogFile = objFSO.OpenTextFile(sPath & "AutoDJ.log", 2, True)
	
' Loop through whole of log file copy and copy relevent lines.
	Do Until objTempLogFile.AtEndOfStream
	
		strLine = objTempLogFile.ReadLine
	
		If Not blnCopy Then
			blnCopy = DateValue(Left(strLine,10)) >= DateValue(Date - intKeepDays)
		End If
		If blnCopy Then objLogFile.WriteLine(strLine)
	Loop

' Close both text files.
	objTempLogFile.Close
	objLogFile.Close

' Delete the copy of the existing log file.
	objFSO.DeleteFile sPath & "AutoDJTemp.log"
	
' Output completed message to the log file.
	OutputToLogFile "Old Log Entries Deleted."
	
End Sub

'*************************************************************************************
' Purpose:	Open the access database and call the macro to export the current list
'			of EAN codes.
'*************************************************************************************
Sub ExportEANs()

	OutputToLogFile "ExportEANs sub run."
	
	RunAccessMacro strAccFullName, "expEANs", 0, False
	
End Sub


'=====================================================================================
'=====================================================================================

'*************************************************************************************
' The following 3 sub procedures correspond to each of the 3 main operation of this
' application (ASN, Aperak, RPO).
'*************************************************************************************

'*************************************************************************************
' Purpose:	To perform all the steps for processing ASN's:
'			1) Move any ASN files from the Shared oneDrive folder to the ASN Folder.
'			2) Check all ASN files in the ASN folder for Priduct Duplicates.
'			3) Merge all ASN files in the ASN Folder in to one CVS file for access.
'			4) Launch the access database and can the Macro for ASN processing.
'			5) After waiting for access to finish, send all approves ASN's via FTP to
'			   david Jones server.
'			6) Move sent files to the archive folder.
'			7) Send an email IF there are any ASN file found to have SSCC numbers that
'			   have already been used.
' Notes:	This sub is to be called a regular Sheduled task as it is to pe run after
'			new ASN files are generated by the AUS warehouse which is currently not
'			something that triggers an event.
'*************************************************************************************
Sub ASNProcessing()
	
	On Error Resume Next
	
	OutputToLogFile "ASNProcessing sub called."
	
' Declare the variables use in this sub procedure.
	Dim fil, strReadAll, strReadAllDups, strReadAllIncorrect, strCurrWorkingFld
	Dim strSearch, arrTempSplit, strFolder, objASNFld, objFil
	Dim objASNDataLoad, strCheckDups, blnSomeASNs, strOneDrivePath
	
' Change the current working Directory to the folder where the new ASN's are.
' But save the previous working directory to a variable, if this is to be changed back
' Later
	strCurrWorkingFld = objShell.CurrentDirectory
	objShell.CurrentDirectory = strASNFolder
	
' Call the function to move multiple files to move any ASN files in the shared onedrive
' folder to the ASN main ASN folder for processing.
' Note: the function will NOT fail or cause an error if there are no files to move.
	strOneDrivePath = objShell.ExpandEnvironmentStrings("%USERPROFILE%") & "\OneDrive - Reiss Limited\ASN\"
	DoCopyMoveFiles strOneDrivePath & "ASNData*.csv", strASNFolder, "MOVE"
	
' Get the ASN folder object.
	set objASNFld = objFso.GetFolder(strASNFolder)

' Initialize the File which all the ASN files are merged to.
' This is done by opening the text file for output, this is to be written to progmatically.
	Set objASNDataLoad = objFSO.OpenTextFile("DJASNImportData.csv", 2, True)

' Loop through each file in the ASN Folder folder.
	For Each objFil In objASNFld.Files
	
	' If the file name contains "ASN" and ".csv" then is it taken to be an ASN file.
		If instr(objFil.Name, "ASNData") > 0 and instr(objFil.Name, ".csv") > 0 Then
		
		' Open the file as a text file for reading only, read all contents to a varible
		' and close the file.
			set fil = objFSO.opentextfile(objFil, 1)
			strReadAll = fil.ReadAll
			fil.Close
		
		' To make things neater, I wrote a separate function for checking a string that is
		' assumed to be the contents of an ASN file for any Product that appear in more
		' than one line and then rectifying it.
		' This function is called for each ASN file, passing in the file contents varaible.
			strCheckDups = CheckForDups(strReadAll)
		
		' The function returns "" if no product dupliates are found. 
			If strCheckDups = "" Then
			
			' If no dups are found write the original ASN file contents to the merg file.
				objASNDataLoad.Write strReadAll
				
				OutputToLogFile objFil.Name & " Appended to data Load, No duplicate Lines found."
				
			Else
			' If duplicates were found the function would have returned the ammended contents
			' of the ASN file, so as well as writing this to the merg file, it aslo needs to
			' be written back to the original ASN file overwritting what was there bofore.
				set fil = objFSO.opentextfile(objFil, 2, True)
				fil.Write strCheckDups
				fil.Close
				objASNDataLoad.Write strCheckDups
				
				OutputToLogFile objFil.Name & " Appended to data Load, Duplicate Lines found."
				
			end If
			
		' This boolean is to indicate that at least one ASN has been found in the ASN folder
		' for processing.
			blnSomeASNs = True
		
		End If
		
' Go to next file in the ASN folder.	
	Next
	
' Close the merge file.
	objASNDataLoad.Close
	
' If at least one ASN file has been processed.
	If blnSomeASNs Then
	
	' Delete the existing Text file uses to output information from acces about any ASN
	' containging SSCC's that have already been used, and Incorrect Unvoice numbers.
	' These files are used for the body of the Response email.
		objFSO.DeleteFile "SSCCAlreadyUsed.txt"
		objFSO.DeleteFile "IncorrectInvNums.txt"
		
	' Open the Access database and run the macro for processing ASN's.
	' Wait for the Access database to close before continuing.
		RunAccessMacro strAccFullName, "processASNs", 0, True
		
	' Send any files that have been placed in the ToSend folder to the david jones servers
	' via a FTP call.
		ExecFTP strASNFolder & "\ToSend", "", "", "", _
				"/", "ASNData*.csv", "mput", False
	
	End If
	
' As access has now processed the information in the merge import file, delete this file.
	objFSO.DeleteFile "DJASNImportData.csv"
	
' If no ASN's were found for processing then the terminate the script.
	If Not blnSomeASNs Then 
		OutputToLogFile "No new ASN's to process."
		Exit Sub
	End If
' Remake the merge file from the files in the ToSend file.
' This new file is to be archived, the reason why it is remade after the Access process gas run
' is to not include any ASN's that were processed but not sent, either because SSCC duplicates
' were found or the RPO is still yet to accepted.
	DoCopyMoveFiles strASNFolder & "\ToSend\ASNData*.csv", strASNFolder & "\DJASNImportData.csv", "COPY"
	
' Move all the sent ASN's to the archive folder.
	DoCopyMoveFiles strASNFolder & "\ToSend\ASNData*.csv", strASNFolder & "\Sent ASNs", "MOVE"

' Copy the new merged file to the archive file prefixing the name wth the DateTime stamp.
	objFSO.CopyFile "DJASNImportData.csv", "Sent ASNs\" & dtDTStamp & "DJASNImportData.csv"

' Check if the file exists that contains information about any ASN's containing duplicae SSCC's
' If there were some dup SSCC's found an email is to be sent out informing people of this.
	If objFSO.FileExists("SSCCAlreadyUsed.txt") Then
		
		OutputToLogFile "SSCC Duplicates Found, see email for details."
		
	' Open the SSCC dup info file, read all contents to a variable and close the file.
		set fil = objFSO.opentextfile("SSCCAlreadyUsed.txt", 1)
		strReadAllDups = fil.ReadAll
		fil.Close
	End If
		
' Check if the file exists that contains information about any incorrect Invoice numbers used.
	If objFSO.FileExists("IncorrectInvNums.txt") Then
		
		OutputToLogFile "Incorrect invoices Found, see email for details"
		
	' Open the Incorect Inv numbers file, read all contents to a variable and close the file.
		set fil = objFSO.opentextfile("IncorrectInvNums.txt", 1)
		strReadAllIncorrect = fil.ReadAll
		fil.Close
	End If

	If Not strReadAllDups = "" Or Not strReadAllIncorrect = "" Then
	
	' Set the whole email response body string incliding the file contents from above.
		strEResponseBody = "Hi All" & vbCrLf & vbCrLf & _
			"Please see below for deatais of ASN errors." & vbCrLf & vbCrLf
		If Not strReadAllDups = "" Then strEResponseBody = strEResponseBody & "SSCC's already used:" & _
			vbCrLf & strReadAllDups & vbCrLf & vbCrLf
		If Not strReadAllIncorrect = "" Then strEResponseBody = strEResponseBody & "Incorrect Invoice numbers/File Name:" & _
			vbCrLf & strReadAllIncorrect & vbCrLf & vbCrLf
		strEResponseBody = strEResponseBody & vbCrLf & vbCrLf & "Best Regards"
			
	' Call the function to assing the email recipient email address to the reciepeint script
	' level variables.
		GetEmailAddresses strEmailsFullName
		
	' Send the email using the custom function for sending an email using the blat application.
		SendBlatEmail strEmailFrom, mstrTo, mstrCC, "", _
				"AutoDJ:ASN Error-Action required", strEResponseBody, False, ""
	End If	
	
	' Move any ASN files that are still in the ASN folder to the ProcessingFailure folder as they
	' would have not been process properly.
	DoCopyMoveFiles strASNFolder & "\ASNData*.csv", strASNFolder & "\ProcessingFailure", "MOVE"		
	
End Sub

'*************************************************************************************
' Purpose:	To perform all the steps required to Process new APERAKS:
'			1) For each file in the Aperak Folder;
'				- Check the file name to deteming if it is an Aperak File.
'				- Check if the aperak has been Accepted.
'				- Check which type af aperak it is (Pricat, Order or Delivery) and
'				  perform required action.
'			2) Open the access database and call the relevnt Macro(s) if any accepted
'			   orders or deliveries were processed.
'			3) Send an email in and deliveries were accepted.
'			4) Send an email if there were any Aperaks that were not accepted.
'			5) If there were any deliveries (RPO's) accepted then move any held ASN's
'			   back for processing again.
' Notes:	This sub is to be called a regular Sheduled task as it is to pe run after
'			new APERAK files are put in the APERAK folder by David Jonnes which is 
'			currently not something that triggers an event.
'*************************************************************************************
Sub AperakProcessing()

' Declare the variables used in this sub.
	Dim fNamePrint, strArchiveFolder, objAperakFolder, objAperakFile, strFolder, objTxtFile
	Dim accOrdersTxt, accDeliveriesTxt, intLen, arrAll, blnAnyNonAccepted, tempVal
	
	On Error Resume Next
	
	OutputToLogFile "AperakProcessing sub Called."
	
' Get the object of the APERAK folder.
	set objAperakFolder = objFSO.GetFolder(strAperakPath)
	
	If Err > 0 Then 
		OutputToLogFile "Error getting object of Aperak Folder"
		Exit Sub
	End If
' Loop through all the files in the APERAK folder.
	For Each objAperakFile In objAperakFolder.Files
	
		If Err > 0 Then Exit Sub
		
	' If the file name has the extension .TXT then it is conciderered to be an Aperak file so
	' continue with the process.
		if instr(objAperakFile.Name, ".TXT") > 0 Then
		
		' Open the file for read only and assing to a variable.
			Set objTxtFile = objFso.OpenTextFile(strAperakPath & "\" & objAperakFile.Name ,1)
		
		' Read the entire contents of the text file and split by line assigning it to an
		' array variable.
			arrAll = Split(objTxtFile.ReadAll, vbCrLf)
		
		' Close the text File.
			objTxtFile.Close
		
		' Initialize the variable used to specify the archive folder that a particular aperak
		' is to be moved to.
			strArchiveFolder = ""
	
		' If "ACCEPTED'" is found in the Aperak the aperak is taken to be accepted.
			if instr(arrAll(5), "ACCEPTED'") > 0 Then
	
			' If the pricat identifier is found set the archive destinatino to the Accepted
			' Pricats folder, this is all that is done for Accepted Pricats.
				if instr(arrAll(3), "PRICAT") then
					strArchiveFolder = "Pricats"
					
					OutputToLogFile "Aperak " & objAperakFile.Name & " - Pricat accepted."
					
			' Else If the order identifier is found, 
				ElseIf instr(arrAll(3), "ORDERS") Then
				
				' Set the archive destination to the Accepte Orders folder.
					strArchiveFolder = "Orders"
					
				' Append the RPO number found in the file text to the sting containing the
				' list of accepetd RPO numbers.
					If not accOrdersTxt = "" then accOrdersTxt = accOrdersTxt & vbNewLine
					tempVal = mid(arrAll(3), instr(arrAll(3), "-") + 1, 7)
					accOrdersTxt = accOrdersTxt & tempVal
					
					OutputToLogFile "Aperak " & objAperakFile.Name & " - RPO " & tempVal & " accepted."
					
			' Else If the delivery identifier is found,
				ElseIf instr(arrAll(3), "DESADV") Then
				
				' Set the archive destinatino to the Accepted Deliveries folder.
					strArchiveFolder = "Deliveries"
					
				' Append the Invoice Numver number found in the file text to the sting containing
				' the list of invoice numbers for all accepted ASN's.
				' Also append to the variable contaning the information for the email response body.
				' Note that the formula below will find an invoice number of varying length, although
				' they will almost always be 6 digits long it mau occasionally contain a M/W suffix
				' if an invoice was processed incorrectily in nova to cantain both men's and womens.
					intLen = instrrev(arrAll(3), "+") - instr(arrAll(3), "-") - 1
					tempVal = mid(arrAll(3), instr(arrAll(3), "-") + 1, intLen)
					strEResponseBody = strEResponseBody & " - inv " & tempVal & vbNewLine
					if Not accDeliveriesTxt = "" then accDeliveriesTxt = accDeliveriesTxt & vbNewLine
					accDeliveriesTxt = accDeliveriesTxt &  tempVal
					
					OutputToLogFile "Aperak " & objAperakFile.Name & " - ASN " & tempVal & " accepted."
					
				End if
			
			' Check if one of the 3 identifiers was found and the archive folder variable was set.
				If not strArchiveFolder = "" Then
				
				' Move the APERAK file to the specified archive folder,
					objFso.moveFile strAperakPath & "\" & objAperakFile.Name, strAperakPath & _
						"\Archive\Accepted\" & strArchiveFolder & "\" & objAperakFile.Name
				
				End If
		
		' Else check If the aperak was NOT accepted.
			Else
			
			' Move the aperak to the folder for Aperaks requiring attention.
				objFso.moveFile strAperakPath & "\" & objAperakFile.Name, strAperakPath & _
					"\AttentionRequired\" & objAperakFile.Name
				
				OutputToLogFile "Aperak " & objAperakFile.Name & " - Attention Required."
				
			' Set the boolean variable indication in at least one Aperak has not been Accepted.
				blnAnyNonAccepted = True
				
			End If
			
		end If

' Move to next Aperak file.
	Next
	
' Check if there were any accepted deliveries, the below variable will be empty if not.
	if Not accDeliveriesTxt = "" Then
	
	' Write the contents of the variable containing the accepted delivery info to a text
	' file to be read by access.
		Set objTxtFile = objFSO.OpenTextFile("Aperak-AcceptedDeliveries.txt", 2, true)
			objTxtFile.write accDeliveriesTxt
			objTxtFile.Close
	
	' Open the Access database and run the macro from ackowledging accepted deliveries.
	' Wait for access to close before continuing the script.
		RunAccessMacro strAccFullName, "importAcceptedDeliveries", 0, True
	
	End If
	
' Check if there were any accepted orders (RPO's), the below variable will be empty if not.
	If Not accOrdersTxt = "" Then
	
	' Write the contents of the variable containing the accepted orders info to a text
	' file to be read by access.
		Set objTxtFile = objFSO.OpenTextFile("Aperak-AcceptedOrders.txt", 2, true)
			objTxtFile.write accOrdersTxt
			objTxtFile.Close
	
	' Open the Access database and run the macro from ackowledging accepted Orders.
	' Wait for access to close before continuing the script.
		RunAccessMacro strAccFullName, "importAcceptedOrders", 0, True
	
	End If
	
' Check if any new deliveries have been accepted by checking the email body variable.
' An email is only sent for Accepted deliveries, not for accepted orders (RPO's) as well.
	if not strEResponseBody = "" Then
	
	' Call the function to assing the email recipient email address to the reciepeint script
	' level variables.
		GetEmailAddresses strEmailsFullName
		
	' Set the whole emaail response body string including the contenst of the strEResponseBody
	' variable.
		strEResponseBody = "Hi All" & vbCrLf & vbCrLf & _
			 "Please note the ASN's for the Following Invoices Have been accepted." & vbCrLf & _
			vbCrLf & strEResponseBody & vbCrLf & _
			"Best Regards" & vbCrLf & vbCrLf
		
	' Send the email using the custom function for sending an email using the blat application.
		SendBlatEmail strEmailFrom, mstrTo, mstrCC, "", _
				"AutoDJ:ASN Acceptance Confirmation", strEResponseBody, False, ""
	
	End If
	
' Check if there are any aperaks that were not accepted.
	If blnAnyNonAccepted Then
	
	' Send the email using the custom function for sending an email using the blat application.
	' The email is to be sent to Admin only as a prompt that some aperaks require some attention.
		SendBlatEmail strEmailFrom, strAdminEmails, "", "", _
				"AutoDJ:Aperak Warning", "There sre some Aperaks that have either been " & _
					"accepted with Error or warning, or Rejected." & vbCrLf & vbCrLf & _
					"They have been moved to Q:\DavidJones\Import\APERAK\AttentionRequired", False, ""
	
	End If
	
' Check again if there were any accepted orders. 
	If Not accOrdersTxt = "" Then
	
	' If there were some accepted Orders processed then the ASN's for these may have already been
	' created and sent over, in which case they will be held in the ASN folder for ASN's where the
	' RPO's have not yet been accepted, so here any ASN files in that folder are moved back to the
	' main ASN folder for processing the next time the ASN process is sheduled to run.
		DoCopyMoveFiles strASNFolder & "\HoldRPONotAccepted\ASNData*.csv", strASNFolder, "MOVE"

	End If
	
End Sub

'*************************************************************************************
' Purpose:	To perform all the steps for Processing a request for RPO's to be created:
'			1) Process the passed string that represents the body of the RPO request
'			   email received, checking it contains any valid requests and preparing
'			   a text file of the invoice numbers accordingly.
'			2) Run the exexcutabe for reprinting invoices and get from Nova server.
'			3) Prepare the raw nova extract in a format for access to process.
'			4) Open access and call the macro for creating RPO's.
'			5) Send the created RPO csv files to the nova server via FTP, and move
'			   files to RPO archive folder.
'			6) Send a response email with info of the RPO's created and any errors.
' Notes:	This sub is generally called by a macro in outlook that is run whenever
'			an email is received that matches the eule criteria for an RPO request. 
' Inputs:	Although there are no arguments suplied drectly to this Sub routine, it
'			uses the "eBody" script input argument. this is the body of the request 
'			email that was received, this is to be passed to the script by the calling
'			outlook macro, however the sub will also handle no argument being passed
'			in the case that the procedure is called manually, this is ecplained below. 
'*************************************************************************************
Sub RpoProcessing()

	On Error Resume Next
	
	OutputToLogFile "RpoProcessing sub Called."
	
' Declare the variables used in this sub.
	Dim eBody, i, strPrepResponse, strEResponseBody, objEResponseTxtF
	Dim blnEmailRec, blnAnyRpoCreated, strManualInputFile, blnRpoAva0
	
' Find if the number of available RPO's is below the warning amount when the process start.
' This is so that the sub can only send an email when the number of available rpo's crosses
' the warning amaount and not each time this is run after.
	blnRpoAva0 = CheckRPOAvailable <= intRPORemaining
	
' The contents of the "eBody" script argument is read in to a variabe.
	eBody = WScript.Arguments.Named.Item("eBody")

' Check if an argument was supplied, If no argument was suplied when running the script
' the variable will be blank.
	If eBody = "fromFile" Then
	
	OutputToLogFile "Email Body to be read from file."
		
	' The following allows for manually running this procedure without outlook calling it.
	' In this case there will be no email body text passed in to the eBody argument, instead
	' the contents of a text file called "ManualRunRPOInput.txt" can be used instead, this
	' file is to lecated in the same folder as this script.
	' First a popup is shown that requires the user to respond, if for some reason this was
	' called progmatically with no argument passed then the popup will not be responded to,
	' in which case it will just close after 6 seconds and the script will terminate accordingly
	' later in the script.
		If Not objShell.Popup("About to run with ManualRunRPOInput.txt file contents, Click Okay to continue.", _
				6, "AutoDJ", 1) = 1 Then 
			OutputToLogFile "User Canceled."
			Exit Sub
		End If
		eBody = ""
		Set strManualInputFile = objFSO.OpenTextFile("ManualRunRPOInput.txt", 1)
		eBody = strManualInputFile.ReadAll
		strManualInputFile.Close
		
	End If
	
' This sub routine requires that these file initially do not exist, and although most or all
' shuold have been Moved or delted on the previous run of this process, if for what ever 
' reason they have not, the will be deleted by the script now.
	objFSO.DeleteFile strImpDataFNm
	objFSO.DeleteFile strRpoErrRepFNm
	objFSO.DeleteFile strRpoEResFNm
	objFSO.DeleteFile strNovaExtRawFNm

' The file delete commands abobove require error handling to be set to resume next as it is
' expected that some files to delete do not exist which returns an error, so next the error
' handling is set back to terminate as i didn't want to leave the who script as resume next.
	On Error Goto 0
	
' Assign the boolean variable that specifies weather or not this process is being run with
' a supplied email body, this includes if it was read from the file after being run manually.
' The reason for this is becasue of the sheduled task that runs for a couple of hours at night
' when Nova is expected to be unavailable, that task call this process with no argument passed
' to run this on the already existing file containing the invoice numbers to be processed.
	blnEmailRec = Not eBody = ""
	
' To make this script more readable there is a separate function for processing a string variable
' representing the body of an email that may contain invoice number for RPO's t be created.
' Below this function is called passing in the email bidy variable.
' The function returns true if there is at least one Invoice to be processed found in the string.
	If CheckInvNums(eBody) Then
		
	' Runt an FTP comman to delete any existing Spoolx Data on the nova server.
	' This is required due to the fact that the spoolx printer apends to an existing file each
	' time which needs to be explicitly deleted, although unlikely, someone else in the complany
	' may have been reprinting invoices to the spoolx printed since the last time this process ran.
		ExecFTP sPath, "", "", "", "", "RTPIIB1S", "", True
		
	' The following loop calls the executable file for reprinting multiple invoices in nova and
	' runs the FTP command to get the SpoolX printer output from the nova server.
	' The reason for this being on a loop is to allow for repeating the process if the executable
	' incorrectle cancelled the printing of an invoice (this may happen due to irregular timing issues
	' with the sending keystrokes process that the executable file runs.
		Do
		
		' Iterate the counter used for the do loop.
			i = i + 1
		
		' Inity
			strPrepResponse = ""
		
		' If it exists already, delete the txt file used to indicate if nova was available or not.
		' This file is genrated by the executable file for reprinting invoices.
		' And delete if exists the nova SpoolX print File.
			On Error Resume Next
			objFSO.DeleteFile("RTPIIB1S")
			On Error Goto 0
		
		' Call the executable file and wait for the file to finish before continuing the script.
			objShell.Run "RePrintMiltipleInvoices.exe", 0, True
			
			OutputToLogFile "RePrintMiltipleInvoices.exe Completed."
			
		' Try runing an FTP command to get the SpoolX output file from the nova server, 
		' this will return false if the file was not found.
		' If a file was retrieved then call the function to prepare the nova extract for
		' processing in access.
			If ExecFTP(sPath, "", "", "", _
					"/", "RTPIIB1S", "get", True) Then _
							strPrepResponse = PrepAccImport(i <= 5)
		
		Loop Until i > 5 Or strPrepResponse = ""
	Else
	
		OutputToLogFile "No new invoice found in Email body for RPO creation."
		
	End If

' Check it the If the access import was prepared succesfully, the function will have returned "Successful".
	If objFSO.FileExists(strImpDataFNm) Then
	
	' Open the access database and call the macro for creating RPO's. Wait for access to close before
	' continuing with the script.
		i = 1
		Do
			On Error Resume Next
			objFSO.DeleteFile("RetryAccess.RpoInfo")
			On Error Goto 0
			
			RunAccessMacro strAccFullName, "createRPOs", 0, True
			i = i + 1
			
		Loop While objFSO.FileExists("RetryAccess.RpoInfo") And i < 4
		
		If objFSO.FileExists("RetryAccess.RpoInfo") Then
		
			objFSO.DeleteFile("RetryAccess.RpoInfo")
			WriteToErrorReport "Error Importing in to Access Tables, no RPO's Processed"
			
		Else
			
		' Run an FTP command to send any RPO's file created by access to the David Jones server.
			ExecFTP sPath, ".", "", "", "/d", "RPO-*.csv", "mput", False
		
		' Move all the RPO csv files that were sent from the export folder to the archive folder.
			DoCopyMoveFiles "RPO-*.csv", strRpoArchivePth, "MOVE"
		
		End If
	
	End If
	
' Assing the boolean variabe for determining if the email response body text file exist, which is oart of
' the logic to determin if an email is to be sent or not.
	blnAnyRpoCreated = objFSO.FileExists(strRpoEResFNm)
	
' An email is to be sent if the email body text file exists OR this process was run with an email received
' body argument. this is to allow for NOT sending an email, if the sheduled task jas called this process
' and there were no RPO's created succesfully.
	If blnAnyRpoCreated Or blnEmailRec Then
	
	' Prepare the start of the email response body.
		strEResponseBody = "Hi All" & vbCrLf & vbCrLf & "Please see below for requested RPO numbers." & vbCrLf & vbCrLf
	
	' Check if there were any RPO's created succesfully.
		If blnAnyRpoCreated Then
		
		' If so, write the contents of the text file to the end of the email response body variable.
			Set objEResponseTxtF = objFSO.OpenTextFile(strRpoEResFNm, 1)
			strEResponseBody = strEResponseBody & objEResponseTxtF.ReadAll
			objEResponseTxtF.Close
		
		' Set the subject to Successful
			strEmailSubject = "Successful"
			
		Else
			
		' Else if no RPO's created bu an email is to be sent, write message to end of the email body variable.
			strEResponseBody = strEResponseBody & "No RPO's have been created." & vbCrLf & vbCrLf & _
				"If the attached error report does not give a clear explanation to why this has happened or " & _
				"it does not exist, please try resending the request email as there may have just " & _
				"been some unexpected interference with the process at the time." & vbCrLf
			OutputToLogFile "#Error# No RPO's Created from request email."
		End If
	
	' Set ending of email body.
		strEResponseBody = strEResponseBody & vbCrLf & "Best Regards"
	
	' If an error report exists, set the sabject according to if any RPO's have been created as well.
		If objFSO.FileExists(strRpoErrRepFNm) Then
		
			If blnAnyRpoCreated Then 
				strEmailSubject = strEmailSubject & " with Error"
			Else
				strEmailSubject = strEmailSubject & "Error"		
			End If
			
		End if
		
		GetEmailAddresses strEmailsFullName
	
	' Send the email using the custom function for sending an email using the blat application.
	' The error report atatchemt file name is always passed, the function for sending the email
	' will check if the file exists and attach if it does but will NOT fail if it does not exist.
		SendBlatEmail strEmailFrom, mstrTo, mstrCC, "", "AutoDJ:RPO Creation-" & strEmailSubject, _
			strEResponseBody, False, strRpoErrRepFNm
	
	' To finish, move any files to be archived to the RPO archive folder prefixing any names with
	' a date and time stamp, and delete the Nova available file that is only need for this particular
	' run of this process.
		On Error Resume Next
		objFSO.MoveFile strImpDataFNm, strRpoArchivePth & dtDTStamp & strImpDataFNm
		objFSO.MoveFile strRpoErrRepFNm, strRpoArchivePth & dtDTStamp & strRpoErrRepFNm
		objFSO.MoveFile strRpoEResFNm, strRpoArchivePth & dtDTStamp & strRpoEResFNm
		objFSO.MoveFile strNovaExtRawFNm, strRpoArchivePth & dtDTStamp & strNovaExtRawFNm
	
	Else
		On Error Resume Next
		objFSO.DeleteFile strImpDataFNm
		objFSO.DeleteFile strRpoErrRepFNm
		objFSO.DeleteFile strRpoEResFNm
		objFSO.DeleteFile strNovaExtRawFNm
		
	End If

' If the number of RPO numbers left is below the threshold and that is different to when the
' sub started, then send the reminder email.
	If blnRpoAva0 <> (CheckRPOAvailable <= intRPORemaining) Then
	
		strEResponseBody = "There are now " & CheckRPOAvailable & " available RPO number's. " & _
			"This is below the warning amount, please request more from David Jones."
			
		SendBlatEmail strEmailFrom, "Tim.Hall@reiss.com,th8313@icloud.com", "", "", "AutoDJ:RPO Numbers Remaing", _
			strEResponseBody, False, ""
			
	End If
	
End Sub

'=====================================================================================
'=====================================================================================

'*************************************************************************************
' Below Are all the functions used by the above script proceedure and sub proceedures.
'*************************************************************************************


'*************************************************************************************
' Purpose:	Reads the text file containing the number if avaiable RPO numbers left.
'*************************************************************************************
Function CheckRPOAvailable()

	On Error Resume Next
	Dim objTxtFile
	Set objTxtFile = objFSO.OpenTextFile(sPath & "RPONumbersAvailable.txt",1)
	CheckRPOAvailable = CInt(Trim(objTxtFile.ReadLine))
	objTxtFile.Close
	
End Function

'*************************************************************************************
' Purpose:	Check ASN files for any product duplications and correct any if found.
' Notes:	ASN Field 20 is the product EAN code
'			ASN Field 36 is the product Quantity
'			ASN Field 38 is the SSCC number	
'*************************************************************************************
Function CheckForDups(strAllTextCSV)

		Dim intLinesDel, arrSplitByLine, arrSplitByField, arrTempSplit
		Dim n, i
		
	' Split the variable by line and assign to an array variable.
		arrSplitByLine = split(strAllTextCSV, vbCrLf)
		
		intLinesDel = 0
		i = 0
		
	' Loop through each line value in the array.
		do until i > ubound(arrSplitByLine) - 1 - intLinesDel

		' If the number of lines deleted so far is greater than 0 then lines will need to
		' be effectivly shifted down so the relevent line is copied to the current array
		' value before processing.
			if intLinesDel > 0 then arrSplitByLine(i) = arrSplitByLine(i + intLinesDel)

			n = 0

		' Split the line value by each comma separated feild and assing to an array variable.			
			arrSplitByField = split(arrSplitByLine(i), ",")

		' Loop through all the line values before the current one.
			Do while n < i
				
			' Split the line values by CSV's.
				arrTempSplit = split(arrSplitByLine(n), ",")

			' Check if the 2 lines contain the same product numerer and SSCC.
				if arrTempSplit(20) = arrSplitByField(20) And _
				arrTempSplit(38) = arrSplitByField(38) then

				' Add the to quantity values together and put in the first line.
					arrTempSplit(36) = cint(arrTempSplit(36)) + cint(arrSplitByField(36))
					
				' Put that whole line back in to the by line Array variable.
					arrSplitByLine(n) = join(arrTempSplit, ",")

				' Increase the lines deleted variable by one.
					intLinesDel = intLinesDel + 1

					i = i - 1
					n = i
				end if
				n = n + 1 
			Loop
			i = i + 1
		Loop
		
	' If some lines have been deleted then...
		if intLinesDel > 0 then
		
		'...redimension the split by line array to delete unwanted values at the end.
			redim preserve arrSplitByLine(ubound(arrSplitByLine) - 1 - intLinesDel)
			CheckForDups = Join(arrSplitByLine, vbCrLf) & vbCrLf
		end If
		
End Function

'*************************************************************************************
' Purpose:		Move or copy files(s).
' Notes:		Unlike the .MoveFile and .CopyFile vbScript function, mutiple files
'				can be moved or copied using wildcards.
' Arguments:	
'*************************************************************************************
Function DoCopyMoveFiles(strSource, strDest, strType)

	On Error Resume Next
	
  	ApplyDblQuotes strSource, True
  	ApplyDblQuotes strDest, True
	
	objShell.Run "%comspec% /c " & strType & " /Y " & strSource & " " & strDest,0,True 

End Function

'*************************************************************************************
' Purpose:	To Write a line of text to the RPO creation error report.
'*************************************************************************************
Function WriteToErrorReport(strText)

	Dim objErrReport
	Set objErrReport = objFSO.OpenTextFile(strRpoErrRepFNm, 8, True)
	objErrReport.WriteLine strText
	objErrReport.Close
	OutputToLogFile "#Error# " & strText
End Function

'*************************************************************************************
' Purpose:	To Write text to the AutoDJ log file.
'			To be used by all sub processes in the AutoDJ System.
'*************************************************************************************
Function OutputToLogFile(strText)
	
	On Error Resume Next
	
	If Right(strText,1) <> "." And Right(strText,1) <> "-" Then strText = strText & "."
	
	Dim objLogFile
	Set objLogFile = objFSO.OpenTextFile(sPath & "AutoDJ.log", 8, True)
	objLogFile.WriteLine Now & " | VBScript: " & strText
	objLogFile.Close
	
End Function

'*************************************************************************************
' Purpose:	Assign values to the mstrTo, mstrCC, mstrBCC script variables.
'			These variables hold to email adresses to send an email to, the address
'			are stored and maintained in the AutoDJEmailRecipients.txt file which
'			should be kept in the same location as this script.
'*************************************************************************************
Function GetEmailAddresses(strFullName)

	on error resume next
	Dim objEmailsTxt, strLine, strSplitLine
	
	' For Testing
	'''strFullName = "I:\Concessions PLU\Data\DavidJones\AutoEDI\AutoDJEmailRecipients_Test.txt"
	
	Set objEmailsTxt = objFSO.OpenTextFile(strFullName, 1)
	
	do until objEmailsTxt.AtEndOfStream
		strLine = objEmailsTxt.ReadLine()
		strSplitLine = split(strLine, ",")
		if Err.Number = 0 then
			if strSplitLine(0) = "TO" then mstrTo = mstrTo & strSplitLine(1) & ","
			if strSplitLine(0) = "CC" then mstrCC = mstrCC & strSplitLine(1) & ","
			if strSplitLine(0) = "BCC" then mstrBCC = mstrBCC & strSplitLine(1) & ","
		end if
		Err.Clear
	Loop
	If Not mstrTo = "" Then mstrTo = Left(mstrTo, Len(mstrTo) - 1)
	If Not mstrCC = "" Then mstrCC = Left(mstrCC, Len(mstrCC) - 1)
	If Not mstrBCC = "" Then mstrBCC = Left(mstrBCC, Len(mstrBCC) - 1)
	
	objEmailsTxt.Close
	
End Function

'*************************************************************************************
' Purpose:	This script reads through the raw invoice data exported from nova and
'			writes it to a new text file in the format required for importing in
'			to the access database to be process for Rpo creation.
'*************************************************************************************
Function PrepAccImport(blnRetyErrors)	

	On Error Resume Next
	
	OutputToLogFile "Preparing Access Import."
	
	dim txtFileRead, txtFileWrite, InvNumber, stNumber, whNum, strLine1, strInvoiceReprint
 	Dim strWholeInvoice, strAllInvoices
 	
 	objFSO.DeleteFile strNovaExtRawFNm
	objFSO.MoveFile "RTPIIB1S", strNovaExtRawFNm
	
' Set the objects of the files to read and to wtrite to the respective variables.	
	
	Set txtFileRead = objFSO.OpenTextFile(strNovaExtRawFNm ,1)

' Loop through each line of the read file.
	blnUsAb = False
	do until txtFileRead.AtEndOfStream

' Set the current line of text to a variable to be read and processed.
		strLine1 = txtFileRead.ReadLine()

' If following 5 if statments check a particular lines text determining if it contains
' required data and processes accordingly.
		If instr(strLine1, "Branch Invoice Number") > 0 Then
			If Not strWholeInvoice = "" Then
				strAllInvoices = strAllInvoices & strWholeInvoice
				strWholeInvoice = ""
			End If
			InvNumber = mid(strLine1, 25, 6)
			
		ElseIf instr(strLine1, " Warehouse ") > 0 then
			whNum = cdbl(trim(mid(strLine1, 12, 8)))
			
		ElseIf instr(strLine1, "Branch    ") > 0 then
			stNumber = cdbl(trim(mid(strLine1, 12, 8)))
			
		ElseIf isNumeric(mid(strLine1,3,6)) Then
			strWholeInvoice = strWholeInvoice & InvNumber & "," & stNumber & "," & whNum & "," & _
			Mid(strLine1, 3, 13) &"," & Trim(mid(strLine1, 17, 25)) & "," & _
			Trim(mid(strLine1, 43, 15)) & "," & Trim(mid(strLine1, 59, 13)) & "," & _
			CDbl(trim(mid(strLine1, 73, 5))) & "," & CDbl(trim(mid(strLine1, 92, 5))) & "," & _
			CDbl(trim(mid(strLine1, 98, 8))) & "," & CDbl(trim(mid(strLine1, 107, 9))) & "," & vbCrLf
			
		ElseIf InStr(strLine1, "LISTING ABORTED BY USER") > 0 Then
			If blnRetyErrors Then
				strInvoiceReprint = strInvoiceReprint & whNum & ":" & InvNumber & vbCrLf
				strWholeInvoice = ""
				OutputToLogFile "Reprinting of Invoice " & InvNumber & " was aborted by user."
			Else
				WriteToErrorReport "The program has repeatedly aborted printing of invoice " & InvNumber
			End If
		End if
	Loop
	
	If Not strWholeInvoice = "" Then strAllInvoices = strAllInvoices & strWholeInvoice
	
' To finish, close the 2 text files
	txtFileRead.Close
	
	If Not strAllInvoices = "" Then
		Set txtFileWrite = objFSO.OpenTextFile(strImpDataFNm, 8, True)
		txtFileWrite.Write strAllInvoices
		txtFileWrite.Close
	End If
	
	objFSO.DeleteFile(strInvNumTxtFname)
	
	If Not strInvoiceReprint = "" Then
		Set txtFileWrite = objFSO.OpenTextFile(strInvNumTxtFname, 2, True)
		txtFileWrite.Write strInvoiceReprint
		txtFileWrite.Close
		PrepAccImport = "Rety"
	End If
	
End Function

'******************************************************************************************
'Purpose:	The followig script reads the data from the EMailBody.txt file line by
'			line, searching for the specified invoice identifier, if found the text
'			following the identifier is writen to the RPOInvoiseNumbers.txt file.
'Assumes:	The EMailBody.txt file has already had the respective emial text outputted
'			to it by Outlook
'******************************************************************************************
Function CheckInvNums(strTextBody)

	dim strPrint, intPos, objTxtFile, strWH, strInv
	dim blnSomeToAppend, strCheckTxt, strInvNumbers
	
	Const strInvIdentifier = "RPOInvoice-"
	
	blnSomeToAppend = False
	if objFSO.fileExists(strInvNumTxtFname) then
		Set objTxtFile = objFSO.OpenTextFile(strInvNumTxtFname, 1)
		strCheckTxt = objTxtFile.ReadAll
		objTxtFile.Close
	end If
	
	intPos = InStr(1, strTextBody, strInvIdentifier)
	Do Until intPos = 0
		strPrint = trim(mid(strTextBody, intPos + len(strInvIdentifier), 9))
		strWH = trim(left(strPrint, 2))
		strInv = trim(mid(strPrint, 4, 6)) 
		if len(strWH) = 2 and isNumeric(strWH) and mid(strPrint, 3, 1) = ":" and _
					len(strInv) = 6 And isNumeric(strInv) then
				if instr(strCheckTxt, strPrint) = 0 then
					strInvNumbers = strInvNumbers & strPrint & vbCrLf
					OutputToLogFile "Invoice " & strPrint & " requested."
				end if
				blnSomeToAppend = True
			else
				WriteToErrorReport "Invoice Identifier found but text following is not correct format"
			End if
		intPos = InStr(intPos + 1, strTextBody, strInvIdentifier)
	Loop
	If Not strInvNumbers = "" Then
		Set objTxtFile = objFSO.OpenTextFile(strInvNumTxtFname, 8, True)
		objTxtFile.Write strInvNumbers
		objTxtFile.Close
	End If
	
	If Not blnSomeToAppend And Not strTextBody = "" Then _
			WriteToErrorReport "No Valid invoices were found in the email"
	CheckInvNums = objFSO.fileExists(strInvNumTxtFname)
	
End Function
'*****************************************************************************************************
' Purpose:	Return a datetime value in the "ddmmyyyy" & "hhmmss" format.
' Input:	dteDateTim is the datetime value to be formatted,
'			blnSep is a boolean value to specify if a seperator is put in between the date and time,
'			a dash "-" will always be used as the seperator.			
'*****************************************************************************************************
Function GetDTStamp(dteDateTime, blnSep)
	Dim strSepValue
	If blnSep Then strSepValue = "-"
	GetDTStamp = Replace(Replace(Replace(dteDateTime, "/", ""), ":", ""), " ", strSepValue)
End Function

'*****************************************************************************************************
' Porpose:	To Open an access database and run a Macro within that database.
' WinStyle:	0 Hide the window (and activate another window.)
' 			1 Activate and display the window. (restore size and position).
' 			2 Activate & minimize. 
' 			3 Activate & maximize. 
' 			4 Restore. The active window remains active. 
' 			5 Activate & Restore. 
' 			6 Minimize & activate the next top-level window in the Z order. 
' 			7 Minimize. The active window remains active. 
' 			8 Display the window in its current state. The active window remains active. 
' 			9 Restore & Activate. Specify this flag when restoring a minimized window. 
'*****************************************************************************************************
Function RunAccessMacro(strAccFName, strMName, intWinStyle, blnWait)

  	ApplyDblQuotes strAccFName, True
  	
  	OutputToLogFile "Access macro called from script."
  	
	objShell.Run "msaccess.exe " & strAccFName & " /x " & strMName, intWinStyle, blnWait
	
End Function	

'*****************************************************************************************************
' Porpose:		To Execute and FTP command.
' Aurguments:	strPLocal - Path on local computer to be used.
'				strSName - Name of the remote server.
'				strSUsername - Username to connect to the remote server.
'				strSPassword - Password to connect to the remote server.
'				strPRemote - Path on the remote server to be used.
'				strFName - The name of the file(s) to me transfered, can contain wildcards.
'				strCmd - A string representing the typ of transfer (i.e. put, get, mput, mget)
'				blnDel - Boolean specifying if the file(s) are to be deleted off the remote server.
'*****************************************************************************************************
Function ExecFTP(strPLocal, strSName, strSUsername, strSPassword, strPRemote, strFName, strCmd, blnDel)
	
	'On Error Resume Next
	' Declare procedure level variables.
	Dim objFTPtxt, objFTPResults
	Dim strPTemp, strTempFile, strFTPResultsFile, strFTPScript, strFTPResults
	Dim strCurrentWorkingFld, strTempWorkingFld

	strPLocal = trim(strPLocal)
	strPRemote = trim(strPRemote)

	If Not objFSO.FolderExists(strPLocal) Then
    	'destination not found
    	ExecFTP = "Error: Local Folder Not Found."
    	Exit Function
  	End If
  
	If InStr(strPRemote, " ") > 0 Then
    		If Left(strPRemote, 1) <> """" And Right(strPRemote, 1) <> """" Then
      			strPRemote = """" & strPRemote & """"
    		End If
  	End If
  	
  	' Get the temp file names.
  	strPTemp = objShell.ExpandEnvironmentStrings("%TEMP%")
	strTempFile = strPTemp & "\" & objFSO.getTempName
	strFTPResultsFile = strPTemp & "\" & objFSO.getTempName
	
	strCurrentWorkingFld = objShell.CurrentDirectory
	objShell.CurrentDirectory = strPLocal

	' Build input file for ftp command.
  	strFTPScript = strFTPScript & "USER " & strSUsername & vbCrLf
  	strFTPScript = strFTPScript & strSPassword & vbCrLf
  	strFTPScript = strFTPScript & "cd " & strPRemote & vbCrLf
  	'strFTPScript = strFTPScript & "binary" & vbCrLf
  	'strFTPScript = strFTPScript & "prompt n" & vbCrLf
  	If Not strCmd = "" Then strFTPScript = strFTPScript & strCmd & " " & strFName & vbCrLf
  	If blnDel Then
  		Dim strDel
  		If Left(strCmd, 1) = "m" Then
  			strDel = "mdelete"
  		Else
  			strDel = "delete"
  		End If
  		strFTPScript = strFTPScript & strDel & " " & strFName & vbCrLf
  	End If
  	strFTPScript = strFTPScript & "quit" & vbCrLf & "quit" & vbCrLf & "quit" & vbCrLf
	
	Set objFTPtxt = objFSO.CreateTextFile(strTempFile, True)
  	objFTPtxt.WriteLine(strFTPScript)
  	objFTPtxt.Close
  	Set objFTPtxt = Nothing  
 	
  	objShell.Run "%comspec% /c FTP -i -n -s:" & strTempFile & " " & strSName & _
  		" > " & strFTPResultsFile, 0, True
	Wscript.Sleep 1000

	objShell.CurrentDirectory = strCurrentWorkingFld
	
	'Check results of transfer.
  	Set objFTPResults = objFSO.OpenTextFile(strFTPResultsFile, 1)
  	strFTPResults = objFTPResults.ReadAll
  	objFTPResults.Close
  	
	' Delete temp files.
	objFSO.DeleteFile(strFTPResultsFile)
	objFSO.DeleteFile(strTempFile)
	
	ExecFTP = InStr(strFTPResults, "226 Transfer complete") > 0
	If ExecFTP Then
		OutputToLogFile "FTP Successful" & vbCrLf & strFTPResults
	Else
		OutputToLogFile "FTP Unsuccessful:" & vbCrLf & strFTPResults
	End If
End Function

'*****************************************************************************************************
' Porpose:	To send an email using the blat application.
' Notes:	The program is run in a command line.	
'*****************************************************************************************************
Function SendBlatEmail(strFrom, strTo, strCC, strBCC, strSubject, strBody, blnFileBody, strFNameAtt)
	
	Dim strBlat
	
	strBlat = "Blat "
	
	If strTo = "" Then Exit Function
	
	ApplyDblQuotes strSubject, False
	ApplyDblQuotes strBody, False
	ApplyDblQuotes strFNameAtt, True
	
	If blnFileBody Then strBlat = strBlat & strBody

	strBlat = strBlat & " -from " & strFrom
	strBlat = strBlat & " -to " & strTo
    If Not strCC = "" Then strBlat = strBlat & " -cc " & strCC
    If Not strBCC = "" Then strBlat = strBlat & " -bcc " & strBCC
    
    strBlat = strBlat & " -subject " & strSubject
    
    If Not blnFileBody Then strBlat = strBlat & " -body " & strBody
    
    If Not strFNameAtt = "" And objFSO.FileExists(strFNameAtt) Then
    
    	strBlat = strBlat & " -attach " & strFNameAtt
    End If
    
    objShell.Run strBlat, 0, True
	
	OutputToLogFile "Email sent, Subject - [ " & strSubject & " ]"
	
End Function

'*****************************************************************************************************
' Porpose:		
' Arguments:	
'*****************************************************************************************************
Function ApplyDblQuotes(ByRef strText, blnCheckSpaces)

	If InStr(strText, " ") > 0 Or Not blnCheckSpaces Then
		If Left(strText, 1) <> """" And Right(strText, 1) <> """" And strText <> "" Then
      		strText = """" & strText & """"
    	End If
    End If
    ApplyDblQuotes = strText
    
End Function

'*****************************************************************************************************
' Porpose:	Check if an AutoDJ process is currently running.
'*****************************************************************************************************
Function IsProcessRunning()
	On Error Resume Next
	IsProcessRunning = objFSO.FileExists(sPath & "ProcessRunning.info")
End Function

'*****************************************************************************************************
' Porpose:	Create the Process running file in the same location as this script to indicate a process
'			is currently running.
'*****************************************************************************************************
Function SetProcessRunning()
	On Error Resume Next
	objFSO.CreateTextFile sPath & "ProcessRunning.info", True
End Function

'*****************************************************************************************************
' Porpose:	Delete the Process running file.	
'*****************************************************************************************************
Function DeleteProcessRunning()
	On Error Resume next
	objFSO.DeleteFile sPath & "ProcessRunning.info"
End Function