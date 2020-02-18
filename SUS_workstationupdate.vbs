'SUSUpdateCodeVersion=1
'   SUSUpdateCodeVersion must be at beginning of line and delimeter is the equal sign with no spaces.
'   Version information is to allow future self-updating of script.
'*******************************************************************************
'*******************************************************************************
' SUS_update.vbs 
' Developed by Mark Colatosti
' Script elements and functions taken from other sources, prodominately Microsoft
' Email function taken from a party on the Internet whose reference has now been lost.
'   It is believed that the email code was public domain as well.
'
'   Creates a Polling Update solution where servers that have this script located
'   in their: 
'   *  c:\windows directory
'   *  The Script Explicitly named "SUS_update.vbs"
'   *  and a scheduled task set to run this script every XXX minutes will result in:
' 
'	1. The ability to download and install updates at a time specified in advance
'      for every server uniquely, specified in a centrally managed "poller" config
'      file located on a read-only web file location. 
'      (recommend co-opting WSUS server if in use)
'	2. The ability to update this script centrally by placing a file named:
'	   SUS_update.vbs.txt in the same web location above and incrementing a
'      version number.  Servers will rewrite their local script file when their 
'      version number is less than the value found in the new file.
'       3. Delay further update attempts after a successful or failed patching for a user
'          specified buffer period, saving IO and CPU overhead of frequent update scans.
'	   Recommend a period of at least one or two days.
'       4. Provide detailed text log file status in the root of the C drive of each 
'	   server for the last patch run.
'	5. If Reboot after patching enabled, a delay, a windows notification balloon
           and scheduled restart is used to minimize impact to a workstation user.
'*******************************************************************************
'*******************************************************************************
' User variables
'*******************************************************************************
' **** HTML Command and Control File Location ****
url = "http://10.129.7.125/machines.txt"
urlUpdate = "http://10.129.7.125/SUS_update.vbs.txt"

'Web address to refer users for unhandled error codes
strAddr = "https://support.microsoft.com/en-us/kb/938205"

'Time period in hours to prevent SUS execution after success or error
' To prevent processing loops!
intLoopGuard = 24

'Filename and location to mark success update (and time)
strUpdateCompletedFilename="c:\WSUS_NoUpdatesRequired"

'Filename and location to mark error in update (and time)
strUpdateErrorFilename="c:\WSUS_UpdatesError"

'Whether or not the user will see the status window.
' Possible options are: 
'0 = verbose, progress indicator, status window, etc.
'1 = silent, no progress indicators.  Everything occurs in the background
Silent = 1 
      
'The location of the logfile (this is the file that will be parsed
' and the contents will be sent via email.                      
'logfile = WshSysEnv("TEMP") & "\" & "vbswsus-status.log"
logfile = "c:\vbswsus-status.log"
									  
'arbitrary email address - reply-to
strMailFrom = "wsus_script@contoso.com"

'who are you mailing to?  Input mode only.  Command-line parameters take 
' precedence
strMailto = "someone@contoso.com"

'set SMTP email server address here
strSMTPServer = "mail.constoso.com"

'set SMTP email server port (default is 25)
iSMTPServerPort = 25

'The computer name will follow this text when the script completes.
strSubject = "[WSUS Update Script] - WSUS Update log file from" 

'Deliminator in above strWUAgentVersion - some locales might have "," instead
' (Non English) - leave as "." if you aren't sure.
strLocaleDelim = "."

'default option for manual run of the script.  Possible options are:
' prompt - (user is prompted to install)
' install - updates will download and install
' detect - updates will download but not install                                       
strAction = "install" 

'Turns email function on/off.  If an email address is specified in the 
' command-line arguments, then this will automatically turn on ('1').
' 0 = off, don't email
' 1 = on, email using default address defined in the var 'strMailto' above.
blnEmail = 0

'strEmailIfAllOK Determines if email always sent or only if updates or reboot 
' needed.
' 0 = off, don't send email if no updates needed and no reboot needed
' 1 = on always send email
strEmailIfAllOK = 0

'strFullDNSName Determines if the email subject contains the full dns name of 
' the server or just the computer name.
' 0 = off, just use computer name
' 1 = on,  use full dns name
strFullDNSName = 0

'tells the script to prompt the user (if running in verbose mode) to input the 
' email address of the recipient they wish to send the script log file to.  The 
' default value in the field is determined by the strMailto variable above.
' 
'This only appears if no command-line arguments are given.  
'0 = do not prompt the user to type in an email address
'1 = prompt user to type in email address to send the log to.
promptemail = 0

'Tells the computer what to do after script execution if the script detects that 
' there is a pending reboot.
'
'Command-prompt supercedes this option.
'0 = do nothing
'1 = reboot
'2 = shutdown
strRestart = 1

'Try to force the script to work through any errors.  Since some are recoverable
' this might be an option for troubleshooting.  Default is 'true'
blnIgnoreError = true

'set your SMTP server authentication type.  
' Possible values:cdoAnonymous|cdoBasic|cdoNTLM
' You do not need to configure an id/pass combo with cdoAnonymous
strAuthType = "cdoAnonymous"

'SMTP authentication ID
strAuthID = ""

'Password for the ID
strAuthPassword = ""

'*******************************************************************************
'End of User variables
'*******************************************************************************

'Handle UAC
If Not WScript.Arguments.Named.Exists("elevate") Then
  CreateObject("Shell.Application").ShellExecute WScript.FullName _
    , WScript.ScriptFullName & " /elevate" & " " & strArguments, "", "runas", 1
  WScript.Quit
End If

Const HKEY_CURRENT_USER 			= &H80000001
Const HKEY_LOCAL_MACHINE 			= &H80000002
Const ForAppending 					= 8
Const ForWriting 					= 2
Const ForReading 					= 1
Const cdoAnonymous 					= 0 'Do not authenticate
Const cdoBasic 						= 1 'basic (clear-text) authentication
Const cdoNTLM 						= 2 'NTLM
Const cdoSendUsingMethod 			= "http://schemas.microsoft.com/cdo/configuration/sendusing", _
			cdoSendUsingPort 		= 2, _
			cdoSMTPServer 			= "http://schemas.microsoft.com/cdo/configuration/smtpserver", _
			cdoSMTPServerport 		= "http://schemas.microsoft.com/cdo/configuration/smtpserverport", _
			cdoSMTPconnectiontimeout = "http://schemas.microsoft.com/cdo/configuration/Connectiontimeout"

On Error Resume Next

' Check if script has an update version, if so update and quit.
Call Reboot(60)
'If VersionUpdateCheck(urlUpdate) = 1 then wscript.quit
' Check if Successfull update has occurred recently
If SUSLoopGuard(intLoopGuard,strUpdateCompletedFilename) = 1 then wscript.quit
' Check if Failed update has occurred recently
If SUSLoopGuard(intLoopGuard,strUpdateErrorFilename) = 1 then wscript.quit

Dim blnRebootRequired
Dim strAction, regWSUSServer, ws, l, wshshell, wshsysenv, strMessage, strFrom
Dim strRestart, silenttext, restarttext, blnCallRestart, blnInstall, blnPrompt, strStatus
Dim blnIgnoreError, blnCScript, strLocaleDelim

Set WshShell = WScript.CreateObject("WScript.Shell")
Set WshSysEnv = WshShell.Environment("PROCESS")
Set ws = wscript.CreateObject("Scripting.FileSystemObject")
Set objADInfo = CreateObject("ADSystemInfo")

' Try to pick up computername via AD'
strComputer1 = objADInfo.ComputerName
' Use the environment strings to pick up the computer name
strComputer = wshShell.ExpandEnvironmentStrings("%Computername%")

' Check Command and Control File
' Determine if computer needs to check for and process updates
intResult = CheckCandC(url,strComputer)

strUser = WshSysEnv("username")
strDomain = WshSysEnv("userdomain")

'Get computer OU
strOU = "Computer OU: Not detected"
Set objComputer = GetObject("LDAP://" & strComputer1)

If objComputer.Parent <> "" Then  
	strOU = "Computer OU: " & replace(objComputer.Parent,"LDAP://","")
End If

If InStr(ucase(WScript.FullName),"CSCRIPT.EXE") Then
	blnCScript = TRUE
Else
	blnCScript = FALSE
End If

blnCloseIE = true

writelog("Log file used: " & logfile)
If intdebug = 1 then wscript.echo "Objargs.count = " & objArgs.count

If blnEmail = 1 and silent = 0 and promptemail = 1 Then strMailto = InputBox("Input the email address you would like the " _
       & "Windows Update agent log sent to:","Email WU Agent logfile to recipient",strMailto)
If strMailto = "" Then wscript.quit

Set l = ws.OpenTextFile (logfile, ForWriting, True)
l.writeline "------------------------------------------------------------------"
l.writeline "WU force update VBScript" & vbcrlf & Now & vbcrlf & "Computer: " & strComputer
l.writeline "Script version: " & strScriptVer
l.writeline strOU 

l.writeline "Executed by: " & strDomain & "\" & strUser
l.writeline "Command arguments: " & strArguments
l.writeline "------------------------------------------------------------------"

If blnEmail = 1 then 
    writelog("SMTP Authentication type specified: " & strAuthType)
    If lcase(strAuthType) <> "cdoanonymous" Then
      If strAuthType = "" Then
        strAuthType = "cdoanonymous"
      Else
        writelog("SMTP Auth User ID: " & sAuthID)
    
        If SMTPUserID = "" then 
          writelog("No SMTP user ID was specified, even though SMTP Authentication was configured for " & strAuthType & ".  Attempting to switch to anonymous authentication...")
          strAuthType = "cdoanonymous"
          If strAuthPassword <> "" then writelog("You have specified a SMTP password, but no user ID has been configured for authentication.  Check the INI file (" & sINI & ") again and re-run the script.")
        Else
          if strAuthPassword = "" then writelog("You have specified a SMTP user ID, but have not specified a password.  Switching to anonymous authentication.")
          strAuthType = "cdoanonymous"
        End if
        If strAuthPassword <> "" then writelog("SMTP password configured, but hidden...")
    
      End If
    End If
End If

Select Case silent
  Case 0
    silenttext = "Verbose"
  Case 1
    silenttext = "Silent"
  Case Else
End Select

If strForceaction = 1 Then 
	strForceText = " (enforce action)"
Else
	strForceText = " (only if action is pending)"
End If

Select Case strRestart
  Case 0
    restarttext = "Do nothing"
  Case 1 
    restarttext = "Restart"
  Case 2 
    restarttext = "Shut down"
  Case Else
End Select

restarttext = restarttext & strForceText

writelog("Script action is set to: " & strAction)
writelog("Verbose/Silent mode is set to: " & silenttext)
writelog("Restart action is set to: " & restarttext)

Set oReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" &_
 strComputer & "\root\default:StdRegProv")
strKeyPath = "SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate"
strValueName = "WUServer"
oReg.GetStringValue HKEY_LOCAL_MACHINE,strKeyPath,strValueName,regWSUSServer
writelog("Checking local WU settings...")

Call GetAUSchedule()

If regWSUSServer then 
Else
  regWSUSServer = "Microsoft Windows Update"
End If

writelog("Update Server: " & regWSUSServer)

strValueName = "TargetGroup"

oReg.GetStringValue HKEY_LOCAL_MACHINE,strKeyPath,strValueName,regTargetGroup
if regTargetGroup <> "" then 
  writelog("Target Group: " & regTargetGroup)
Else
  writelog("Target Group: Not specified")
End If

Set autoUpdateClient = CreateObject("Microsoft.Update.AutoUpdate")
Set updateInfo = autoUpdateClient.Settings

Select Case updateInfo.notificationlevel
	Case 0 
	  writelog("WUA mode: WU agent is not configured.")
	Case 1 
	  writelog("WUA mode: WU agent is disabled.")
	Case 2
	  writelog("WUA mode: Users are prompted to approve updates prior to installing")
	Case 3 
	  writelog("WUA mode Updates are downloaded automatically, and users are prompted to install.")
	Case 4 
	  writelog("WUA mode: Updates are downloaded and installed automatically at a predetermined time.")
	Case Else
End Select

On Error Resume Next

Set updateSession = CreateObject("Microsoft.Update.Session")
Set updateSearcher = updateSession.CreateupdateSearcher()

writelog("Instantiating Searcher")
Set searchResult = updateSearcher.Search("IsAssigned=1 and IsHidden=0 and IsInstalled=0 and Type='Software'")

'Handle some common errors here
If cstr(err.number) <> 0 Then
  If cstr(err.number) = "-2147012744" Then
    strMsg = "ERROR_HTTP_INVALID_SERVER_RESPONSE - The server response could not be parsed." & vbcrlf & vbcrlf & "Actual error was: " _
      & " - Error [" & cstr(err.number) & "] - '" & err.description & "'"
    blnFatal = true
  ElseIf CStr(err.number) = "-2145107924" Then
    strMsg = "WU_E_PT_WINHTTP_NAME_NOT_RESOLVED - Winhttp SendRequest/ReceiveResponse failed with 0x2ee7 error. Either the proxy " _
     & "server or target server name can not be resolved. Corresponding to ERROR_WINHTTP_NAME_NOT_RESOLVED. " _
     & "Stop/Restart service or reboot the machine if you see this error frequently. " _
     & vbcrlf & vbcrlf & "Actual error was [" & err.number & "] - " & chr(34) _
      & err.description & chr(34)
    blnFatal = false
  ElseIf cstr(err.number) <> 0 and cstr(err.number) = "-2147012867" Then 
    strMsg = "ERROR_INTERNET_CANNOT_CONNECT - The attempt to connect to the server failed." & vbcrlf _
      & vbcrlf & "Actual error was [" & err.number & "] - " & chr(34) _
      & err.description & chr(34)
    blnFatal = true
  ElseIf CStr(err.number) = "-2145107941" Then 
    strMsg = "SUS_E_PT_HTTP_STATUS_PROXY_AUTH_REQ - Http status 407 - proxy authentication required" & vbcrlf & vbcrlf & "Actual " _
     & "error was [" & err.number & "]" & chr(34) & err.description & chr(34)
  ElseIf CStr(err.number) = "-2145124309" Then 
    strMsg = "WU_E_LEGACYSERVER - The Sus server we are talking to is a Legacy Sus Server (Sus Server 1.0)" _
     & vbcrlf & vbcrlf & "Actual error was [" & err.number & "] - " & chr(34) & err.description & chr(34)
    blnFatal = true
  ElseIf CStr(err.number) = "7" Then 
    strMsg = "Out of memory - In most cases, this error will be resolved by rebooting the client." _ 
     & VbCrLf & VbCrLf & "Actual error was [" & err.number & "] - " & chr(34) & err.description & chr(34) 
    blnFatal = True 
  Else
    If err.description = "" Then 
    	errdescription = "No error description given"
    Else 
        errdescription = err.description
    End If
    If blnIgnoreError = false Then 
    	blnFatal = true 
 	    strScriptAbort = vbcrlf & vbcrlf & "Script will now abort. - if you want to force the script to continue, change the 'blnIgnoreError' variable " _
     	 & "to the value 'true'"
    Else
    	strScriptabort = vbcrlf & vbcrlf & "Script will attempt to continue."
    End If
    
    strMsg = "Error - [" & err.number & "] - " & chr(34) & errdescription & chr(34) & "." & vbcrlf & vbcrlf _
     & "This error is undefined in the script, but you can refer to " & strAddr & " to look up the error number." _
     & strScriptAbort
     strMsgHTML = replace(strMsg,strAddr,"<a href='" & strAddr & "'>" & strAddr & "</a>")
    If silent = 0 Then objdiv.innerhtml = replace(strMsgHTML,"vbcrlf","<br>")
   End If
  
  Call ErrorHandler("UpdateSearcher",strMsg,blnFatal)
End If

Call CheckPendingStatus("beginning")

strMsg = "Refreshing WUA client information..."

'cause WU agent to detect
on error resume next
autoUpdateClient.detectnow()
if err.number <> 0 then call ErrorHandler("WUA refresh",err.number & " - " & err.description,false)
err.clear
on error goto 0 

writelog("WUA mode: " & straction)
writelog("WU Server: " & regWSUSServer)
writelog("Searching for missing or updates not yet applied...")

on error resume next

writelog("Missing " & searchResult.Updates.Count & " update(s).") 

For i = 0 To searchResult.Updates.Count-1
    Set update = searchResult.Updates.Item(i)
	strSearchResultUpdates = strSearchResultUpdates & update.Title & "<br>"
    Set objCategories = searchResult.Updates.Item(i).Categories
    writelog("Missing: " & searchResult.Updates.Item(i))
Next


if err.number <> 0 then
    writelog("An error has occured while instantiating search results.  Error " & err.number & " - " & err.description _
        & ".  Check the " & wshShell.ExpandEnvironmentStrings("%windir%") & "\windowsupdate.log file for further information.")
    blnFatal = false
End IF
    
If searchResult.Updates.Count = 0 Then
  Set objShell = WScript.CreateObject ("WScript.shell")
  'Report status to SUS server
  objShell.run "c:\windows\system32\wuauclt.exe /reportnow"
  Set objShell = Nothing
  'No Updates needed - Update file marking last no update check date/time
  ' This file prevents further updates for a specific buffer time.
  '~ Create a FileSystemObject
	Set objNoUpdate=CreateObject("Scripting.FileSystemObject")
  '~ Setting up file to write
	Set objFileNoUpdate = objNoUpdate.CreateTextFile(strUpdateCompletedFilename,True)
    objFileNoUpdate.WriteLine Now
  '~ Close the file
	objFileNoUpdate.Close
	
  strMsg = fformat & "There are no further updates needed for your PC at this time."
  writelog(replace(strMsg,fformat,""))
  writelog("Events saved to '" & logfile & "'")
  
  Call EndOfScript
  wscript.quit
End If

If intdebug = 1 then WScript.Echo vbCRLF & "Creating collection of updates to download:"
If strAction <> "detect" Then writelog("Creating a catalog of needed updates") 

writelog("********** Cataloging updates **********")

Set updatesToDownload = CreateObject("Microsoft.Update.UpdateColl")

For I = 0 to searchResult.Updates.Count-1 
	Set update = searchResult.Updates.Item(I) 
	if update.MsrcSeverity <> "" then MsrcSeverity = "(" & update.MsrcSeverity & ") "
	strUpdates = strUpdates & MsrcSeverity & "- " & update.Title & "<br>"
 	writelog("Cataloged: " & MsrcSeverity & update.Title) 
	If Not update.EulaAccepted Then update.AcceptEula 
	updatesToDownload.Add(update) 
Next 

strMsg = fformat & "This PC requires updates from the configured Update Server" _
 & " (" & regWSUSServer & ").  "
If strAction <> "detect" Then strmsg = strmsg & "<br><br> Downloading needed updates.  Please stand by..."
writelog(replace(replace(strMsg,fformat,""),"<br>",""))

If strAction = "detect" Then 
	
Else
	
	Set downloader = updateSession.CreateUpdateDownloader() 
	on error resume next
	downloader.Updates = updatesToDownload
	writelog("********** Downloading updates **********")

	downloader.Download()

	if err.number <> 0 then
		writelog("Error " & err.number & " has occured.  Error description: " & err.description)
	End if

	strUpdates = ""
	strMsg = ""
	strMsg = fformat & "List of downloaded updates: <br><br>"
	
	For I = 0 To searchResult.Updates.Count-1
	    Set update = searchResult.Updates.Item(I)
	    If update.IsDownloaded Then
	       strDownloadedUpdates = strDownloadedUpdates & update.Title & "<br>"
	    End If
	       On Error GoTo 0
	       'writelog(searchResult.Updates.Item(i))
	       writelog("Downloaded: " & update.Title)
	Next
	Set updatesToInstall = CreateObject("Microsoft.Update.UpdateColl")
	
	strUpdates = ""
	strMsg = ""
	strMsg = fformat & "Creating collection of updates needed to install:<br><br>" 
	writelog("********** Adding updates to collection **********")

	For I = 0 To searchResult.Updates.Count-1
	    set update = searchResult.Updates.Item(I)
	    If update.IsDownloaded = true Then
	       strUpdates = strUpdates & update.Title & "<br>"
	       updatesToInstall.Add(update)
	    End If
	       writelog("Adding to collection: " & update.Title)
	Next
End If


If lcase(strAction) = "prompt" Then 
  strMsg = "The Windows Update Agent has detected that this computer is missing updates from the " _
   & " configured server (" & regWSUSServer & ")." & vbcrlf & vbcrlf & "Would you like to install updates now?"
  strResult = MsgBox(strMsg,36,"Install now?")
  strUpdates = ""
  writelog(strMsg & " [Response: " & strResult & "]")
  strMsg = ""
ElseIf strAction = "detect" Then
  strMsg = fformat & "Windows Update Agent has finished detecting needed updates." 
  writelog(replace(strMsg,fformat,""))
  Call EndOfScript
  wscript.quit
ElseIf strAction = "install" Then
  strResult = 6
End If 

strUpdates = ""

If strResult = 7 Then
  strMsg = strMsg & "<br>User cancelled installation.  This window can be closed."
  writelog(replace(strMsg,"<br>",""))
	WScript.Quit
ElseIf strResult = 6 Then
  strMsg = ""
  Set installer = updateSession.CreateUpdateInstaller()
  installer.AllowSourcePrompts = False 
  on error resume next 

  installer.ForceQuiet = True 
  
  strMsg = fformat & "Installing updates... <br><br>"
  writelog(replace(replace(strMsg,fformat,""),"<br>",""))
  
 If err.number <> 0 Then
	writelog("Error " & err.number & " has occured.  Error description: " & err.description)
 End if
  
	installer.Updates = updatesToInstall
	
	writelog("********** Installing updates **********")
	
	blnInstall = true
	
	on error resume next	
	Set installationResult = installer.Install()

	writelog(replace(replace(strMsg,fformat,""),"<br>",""))
 	
	If err.number <> 0 then 
	    strMsg = "Error installing updates... Actual error was " & err.number & " - " & err.description & "."
	    writelog(strmsg)
		'Updates Failure - Update file marking an update failure occurred
		' This file prevents further updates for a specific buffer time.
		'~ Create a FileSystemObject
		Set objNoUpdate=CreateObject("Scripting.FileSystemObject")
		'~ Setting up file to write
		Set objFileNoUpdate = objNoUpdate.CreateTextFile(strUpdateErrorFilename,True)
		objFileNoUpdate.WriteLine Now
		'~ Close the file
		objFileNoUpdate.Close
	End If

	'Output results of install
	strMsg = fformat & "Installation Result: " & installationResult.ResultCode & "<br><br>" _
	 & "Reboot Required: " & installationResult.RebootRequired & "<br><br>" _
	 & "Listing of updates and individual installation results: <br>"

	 For i = 0 to updatesToInstall.Count - 1
		WriteLog("Installing " & update.title)
 
		If installationResult.GetUpdateResult(i).ResultCode = 2 Then 
			strResult = "Installed"
		ElseIf installationResult.GetUpdateResult(i).ResultCode = 1 Then 
			strResult = "In progress"
		ElseIf installationResult.GetUpdateResult(i).ResultCode = 3 Then 
			strResult = "Error"
		ElseIf installationResult.GetUpdateResult(i).ResultCode = 4 Then 
			strResult = "Failed"
		ElseIf installationResult.GetUpdateResult(i).ResultCode = 5 Then 
			strResult = "Aborted"			
		End If
		writelog(updatesToInstall.Item(i).Title & ": " & strResult)
    strUpdates = strUpdates & strResult & ": " & updatesToInstall.Item(i).Title & "<br>"
	Next
End If		

Call EndOfScript
wscript.quit

'*******************************************************************************
'Function Writelog 
'*******************************************************************************
Function WriteLog(strMsg) 
l.writeline "[" & time & "] - " & strMsg
' Output to screen if cscript.exe 
If blnCScript Then WScript.Echo "[" & time & "] " & strMsg 
End Function 

Sub GetAUSchedule()
Set objAutoUpdate = CreateObject("Microsoft.Update.AutoUpdate")
Set objSettings = objAutoUpdate.Settings

Select Case objSettings.ScheduledInstallationDay
    Case 0
        strDay = "every day"
    Case 1
        strDay = "sunday"
    Case 2
        strDay = "monday"
    Case 3
        strDay = "tuesday"
    Case 4
        strDay = "wednesday"
    Case 5
        strDay = "thursday"
    Case 6
        strDay = "friday"
    Case 7
        strDay = "saturday"
    Case Else
        strDay = "The scheduled installation day is could not be determined."
End Select

If objSettings.ScheduledInstallationTime = 0 Then
    strScheduledTime = "12:00 AM"
ElseIf objSettings.ScheduledInstallationTime = 12 Then
    strScheduledTime = "12:00 PM"
Else
    If objSettings.ScheduledInstallationTime > 12 Then
        intScheduledTime = objSettings.ScheduledInstallationTime - 12
        strScheduledTime = intScheduledTime & ":00 PM"
    Else
        strScheduledTime = objSettings.ScheduledInstallationTime & ":00 AM"
    End If
    'strTime = "Scheduled installation time: " & strScheduledTime
End If

writelog("Windows update agent is scheduled to run on " & strDay & " at " & strScheduledTime)
End Sub

'*******************************************************************************
'Function SendMail - email the warning file
'*******************************************************************************
Function SendMail(strFrom,strTo,strSubject,strMessage)
Dim iMsg, iConf, Flds

writelog("Calling sendmail routine")
writelog("To: " & strMailto)
writelog("From: " & strMailFrom)
writelog("Subject: " & strSubject)
writelog("SMTP Server: " & strSMTPServer)

'//  Create the CDO connections.
Set iMsg = CreateObject("CDO.Message")
Set iConf = CreateObject("CDO.Configuration")
Set Flds = iConf.Fields

If lcase(strAuthType) <> "cdoanonymous" Then
  'Type of authentication, NONE, Basic (Base64 encoded), NTLM
  iMsg.Configuration.Fields.Item _
("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = strAuthType

  'Your UserID on the SMTP server
  iMsg.Configuration.Fields.Item _
("http://schemas.microsoft.com/cdo/configuration/sendusername") = strAuthID

  'Your password on the SMTP server
  iMsg.Configuration.Fields.Item _
("http://schemas.microsoft.com/cdo/configuration/sendpassword") = strAuthPassword

End if

'// SMTP server configuration.
With Flds
	.Item(cdoSendUsingMethod) = cdoSendUsingPort
	.Item(cdoSMTPServer) = strSMTPServer
	.Item(cdoSMTPServerPort) = iSMTPServerPort
	.Item(cdoSMTPconnectiontimeout) = 60
	.Update
End With
'l.close

Dim r
Set r = ws.OpenTextFile (logfile, ForReading, False, TristateUseDefault)
strMessage = "<font face='" & strFontStyle & "' size='2'>" & r.readall & "</font>"

'//  Set the message properties.
With iMsg
    Set .Configuration = iConf
        .To       = strMailTo
        .From     = strMailFrom
        .Subject  = strSubject
        '.TextBody = strMessage
End With

'iMsg.AddAttachment wsuslog
iMsg.HTMLBody = replace(strMessage,vbnewline,"<br>")
'//  Send the message.
on error resume next

iMsg.Send ' send the message.
Set iMsg = nothing

If CStr(err.number) <> 0 Then
	strMsg = "Problem sending mail to " & strSMTPServer & "." _
   & "Error [" & err.number & "]: " & err.description & "<br>"
  
  Call ErrorHandler("Sendmail function",replace(strMsg,"<br>",""),"false")
  'writelog(strMsg)
  strStatus = strMsg
Else
  strStatus = "Connected successfully to email server " & strSMTPServer
  writelog(strStatus)
	strStatus = strStatus & "<br><br><font face=" & strFontStyle & " color=" & strFontColor2& ">" _
 & "sent email to " & strMailTo & "...</font><br><BR>" _
	 & "Script complete.<br><br><a href='file:///" & logfile & "'>View log file</a>"
End If

'cause WU agent to detect
autoUpdateClient.detectnow()
blnEmail = 0

End Function
'*******************************************************************************'Function RestartAction
'Sub to perform a restart action against the computer
'*******************************************************************************
Function RestartAction
  wscript.sleep 4000
  writelog("Processing PostExecuteAction")
	'On Error GoTo 0
	Dim OpSysSet, OpSys
	'Call WMI query to collect parameters for reboot action
	Set OpSysSet = GetObject("winmgmts:{(Shutdown)}//" & strComputer & "/root/cimv2").ExecQuery("select * from Win32_OperatingSystem"_
	 & " where Primary=true") 
	 
	If CStr(err.number) <> 0 Then 
	  strMsg = "There was an error while attempting to connect to " & strComputer & "." & vbcrlf & vbcrlf _
		 & "The actual error was: " & err.description
		writelog(strMsg)
		blnFatal = true
    	Call ErrorHandler("WMI Connect",strMsg,blnFatal)
	End If

  	Const EWX_LOGOFF = 0 
  	Const EWX_SHUTDOWN = 1 
  	Const EWX_REBOOT = 2 
  	Const EWX_FORCE = 4 
  	Const EWX_POWEROFF = 8 
	
	'set PC to reboot
	If strRestart = 1 Then

		'For each OpSys in OpSysSet 
		'	opSys.win32shutdown EWX_REBOOT + EWX_FORCE
		'Next
		Call Reboot(30) 

	'set PC to shutdown
	ElseIf strRestart = 2 Then
				
		For each OpSys in OpSysSet 
			opSys.win32shutdown EWX_POWEROFF + EWX_FORCE
		Next 
  
  'Do nothing...
  ElseIf strRestart = "0" Then
    				
End If


End Function

'*******************************************************************************
'Sub ErrorHandler
'Sub to help display/log any errors that occur
'*******************************************************************************
Sub ErrorHandler(strSource,strMsg,blnFatal)
    'Set theError = RemoteScript.Error
		writelog(strMsg)
		If blnFatal = true then wscript.quit
    err.clear
End Sub

'*******************************************************************************'Function EndOfScript
'Function to close out the script
'*******************************************************************************
Function EndOfScript
  Dim objShell
  Set objShell = WScript.CreateObject ("WScript.shell")
  'Report status to SUS server
  objShell.run "c:\windows\system32\wuauclt.exe /reportnow"
  Set objShell = Nothing
  If blnInstall = true then Call CheckPendingStatus("end")
  on error goto 0
  writelog("Windows Update VB Script finished")
  l.writeline "---------------------------------------------------------------------------"
  If blnCallRestart = true then writelog("Post-execute action will be called.  " _
   & " Action is set to: " & restarttext & ".")
     

  If blnEmail = 1 Then
     If searchresult.updates.count = 0 and not blnRebootRequired and StrEmailifAllOK = 0 then
        writelog ("No updates required, no pending reboot, therefore not sending email")
     else
        if strFullDNSName = 1 then
           strDomainName = wshShell.ExpandEnvironmentStrings("%USERDNSDOMAIN%")
	  			 strOutputComputerName = strComputer & "." & StrDomainName
        else
           strOutputComputerName = strComputer         
        end if
        if emailifallok = 0 or emailifallok = 1 then
          if instr(strSMTPServer,"x") then
          else
           Call SendMail(strFrom,strTo,strSubject & " " & strOutputComputerName,strMessage)
          end if
        end if
     end if
  Else
  End If

  strMsg = "The script has been configured to " & restarttext _
   		& ".  The update script has detected that this " _
   		& "computer has a reboot pending from a previous update session." & vbcrlf & vbcrlf _
   		& "Would you like to perform this action now?"
  
  If silent = 0 and blnPrompt = true Then 
  	strResult = MsgBox(strMsg,36,"Perform restart/shutdown action?")
  ElseIf blnPrompt = false Then
  	strResult = 6
  End If
     
  If blnCallRestart = true Then 
  	If strResult = 6 Then call RestartAction
  Else
    on error resume next
  End If
  wscript.quit
  
  Exit Function
   
End Function

'*******************************************************************************
'Function CheckPendingStatus
'Function to restart the computer if there is a reboot pending...
'*******************************************************************************
Function CheckPendingStatus(beforeorafter)
  Set ComputerStatus = CreateObject("Microsoft.Update.SystemInfo")
  Select case beforeorafter
    Case "beginning"
      strCheck = "Pre-check"
    Case "end"
      strCheck = "Post-check"
    Case Else
  End Select
  
  blnRebootRequired = ComputerStatus.RebootRequired

  If ComputerStatus.RebootRequired or strForceAction = 1 Then
     If beforeorafter = "beginning" Then 
        If ComputerStatus.RebootRequired Then strMsg = "This computer has a pending reboot (" & strCheck & ").  Switching to 'detect' mode."
        If strAction = "prompt" Then blnPrompt = true
        strAction = "detect"
        blnCallRestart = true  
     Else
        If ComputerStatus.RebootRequired Then strMsg = "This computer has a pending reboot (" & strCheck & ").  Setting PC to perform post-script " _
          & "execution..."
        blnCallRestart = true        
     End If
  Else
        If not ComputerStatus.RebootRequired Then strMsg = "This computer does not have any pending reboots (" & strCheck & ")."
  End If
  
     If strMsg <> "" Then writelog(strMsg)
     'wscript.sleep 4000
           
End Function
'*******************************************************************************

'*******************************************************************************
'Function CheckCandC
'Function to read web-based c&c file
'*******************************************************************************
Function CheckCandC(url,strComputer)
	on error resume next
	Set xmlhttp = CreateObject("MSXML2.ServerXMLHTTP")
	xmlhttp.open "GET", url, 0
	xmlhttp.send ""
	' If error <>0 then we were not able to retrieve C&C file!
	if err.number<>0 then WScript.Quit
	
	strHTTP = xmlhttp.responseText
	Set xmlhttp = Nothing

	' Will be used to determine if updates are scheduled to run
	strDateTime = Now
	' Takes HTTP data and splits into an array of separate line items
	strHTTPSplit=Split(strHTTP,vbCrlf)
	' Set default action of update NOT to occur
	boolUpdate=false

	for each strLine in strHTTPSplit
	' Split HTTP GET lines into individual server entries
		arrServerList = Split(strLine, ",") 
		strServer = arrServerList(0) 
		strUpdateTime = arrServerList(1)
		' Check to see if this server is in the downloaded list
		if (lcase(strServer) = lcase(strComputer) Or lcase(strServer) = "all") then 
		' If retrieved update date/time - NOW is negative, then updates are schuduled to be applied.
			if DateDiff("s",strDateTime,strUpdateTime) < 0 then 
				boolUpdate=true
			end if
		end if
	Next
	' CRITICAL CHECK
	' If server name not in published list or not the time to update, exit and do not install updates!
	If boolUpdate = false Then WScript.Quit
End Function

Function SUSLoopGuard(intLoopGuard,strOutFile)
' Check if system recently updated completely, if so exit immediately!
	On error resume next
	Set objNoUpdate=CreateObject("Scripting.FileSystemObject")
	'~ Setting up file to write
	Set objFileNoUpdate = objNoUpdate.GetFile(strOutFile)
	intLastRun = 9999999
	intLastRun = datediff("h",objFileNoUpdate.DateCreated,Now)
	if intLastRun < intLoopGuard then
		objFileNoUpdate.Close
		' Updates have occurred recently do not process again
		SUSLoopGuard = 1
	else
		objNoUpdate.DeleteFile strOutFile
		SUSLoopGuard = 0
	end if
	objFileNoUpdate.Close
End Function

Function VersionUpdateCheck(url)
	on error resume next
	Set xmlhttp = CreateObject("MSXML2.ServerXMLHTTP")
	xmlhttp.open "GET", url, 0
	xmlhttp.send ""
	' If error <>0 then we were not able to retrieve C&C file!
	if err.number<>0 then VersionUpdateCheck=0
	strHTTP = xmlhttp.responseText
	Set xmlhttp = Nothing
	
	Set objRegEx = CreateObject("VBScript.RegExp")
	' Regex pattern should only match the version string
	objRegEx.Pattern = "^'SUSUpdateCodeVersion=(\d*)"
	objRegEx.Multiline = true
	objRegEx.IgnoreCase = true
	Set RegexMatches = objRegEx.Execute(strHTTP)
	If Regexmatches.Count > 0 Then
		Set match = Regexmatches(0)
		arrUpdateVersionList = Split(match.Value, "=") 
		strUpdateVersiontxt = arrUpdateVersionList(0) 
		strUpdateVersionNumber = arrUpdateVersionList(1)
	Else
		VersionUpdateCheck=0
	End if
	
	'Read existing SUS_update.vbs file version
	Const ForReading = 1
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Set objVBSFile = objFSO.OpenTextFile("C:\windows\SUS_update.vbs", ForReading)
	if err.number<>0 then VersionUpdateCheck=0
	strVersion = objVBSFile.ReadLine
	if err.number<>0 then VersionUpdateCheck=0
	objVBSFile.Close
	arrCurrentVersionList = Split(strVersion, "=")	
	strCurrentVersiontxt = arrCurrentVersionList(0) 
	strCurrentVersionNumber = arrCurrentVersionList(1)

	if (int(strCurrentVersionNumber) < int(strUpdateVersionNumber)) then
		Set objVBSFile = objFSO.CreateTextFile("c:\windows\SUS_update.vbs",True)
		objVBSFile.Write strHTTP
		objVBSFile.Close
		'Quit script after updating file so that it will run on next cycle instead of old script being allowed to run now.
		VersionUpdateCheck=1
	else
		VersionUpdateCheck=0
	end if
End Function


Function Reboot(duration)
	'Option Explicit
	Dim fso,Ws,Ret,ByPassPSFile,PSFile
	PSFile = "C:\BallonNotification" & "." & "ps1"
	Set Ws = CreateObject("wscript.Shell")
	ByPassPSFile = "cmd /c PowerShell.exe -ExecutionPolicy bypass -noprofile -file "
	displayText = "'Your computer will restart at " & FormatDateTime(dateadd("n",duration,Now),4) & " ! Please save your work immediately and manually restart your computer.'"
	Call WritePSFile("Warning","60","'Security Updates Require System Restart'",displayText,"'Warning'","60")
	Ret = Ws.run(ByPassPSFile & PSFile,0,True)
	Set fso = CreateObject("Scripting.FileSystemObject")
	If fso.FileExists(PSFile) Then
		fso.DeleteFile(PSFile)
	End If
	durationSeconds = cstr(cint(duration) * 60)
	shutdownCommand = "cmd /c shutdown /r /t " & durationSeconds & " /c ""Security updates require a system restart!"" /d p:2:18"
	Ret = Ws.run(shutdownCommand,0,True)
End Function
'------------------------------------------------------------------------------------------------------------
Sub WritePSFile(notifyicon,time,title,text,icon,Timeout) 
	Const ForWriting = 2
	Dim fso,ts,strText
	PSFile = "C:\BallonNotification" & "." & "ps1"
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set ts = fso.OpenTextFile(PSFile,ForWriting,True)
	strText = strText & "[reflection.assembly]::loadwithpartialname('System.Windows.Forms') | Out-Null;" & VbCrlF
	strText = strText & "[reflection.assembly]::loadwithpartialname('System.Drawing') | Out-Null;" & VbCrlF 
	strText = strText & "$notify = new-object system.windows.forms.notifyicon;" & VbCrlF
	strText = strText & "$notify.icon = [System.Drawing.SystemIcons]::"& notifyicon &";" & VbCrlF 
	strText = strText & "$notify.visible = $true;" 
	strText = strText & "$notify.showballoontip("& 0 &","& title &","& text &","& icon &");" & VbCrlF 
	'strText = strText & "Start-Sleep -s " & Timeout &";" & VbCrlF
	'strText = strText & "$notify.Dispose()"
	ts.WriteLine strText
End Sub