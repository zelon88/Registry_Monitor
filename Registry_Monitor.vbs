'File Name: Registry_Monitor.vbs
'Version: v1.0, 9/19/2019
'Author: Justin Grimes, 8/29/2019

'Supported Arguments
  ' -e  (Email)  =  Set "emailResult" config entry to TRUE (send emails when registry changes are detected).
  ' -o  (Output)  =  Set "outputResult" config entry to TRUE (create a log file when registry changes are detected).
  ' -v  (verbose)  =  Set "verbose" config entry to TRUE (create a log file whenever the script is executed).
  ' -f  (Forced)  =  Set "force" config entry to TRUE (bypass elevated priviledge checks, may encounter errors).
' --------------------------------------------------

' --------------------------------------------------
'Declare all variables to be used during execution of this script.
'Undeclared variables will cause a critical error and halt script execution.
Option Explicit
Dim strKeyPath, hive, hiveItem, key, reg, arrSubKeys, subkey, objFSO, regFileHandle1, regFilePath1, regFileHandle2, regFilePath2, hiveArray, objShell, _
 objScript, scriptPath, cachePath,  RNscriptName, RNappPath, RNlogPath, companyName, companyAbbr, companyDomain, toEmail, RNmailFile, logFilePath, Path, _
 reqdDirsExists, newRegData, outputResult, strSafeDate, strSafeTime, strDateTime, logFileName, strComputerName, oShell2, strUserName, _
 tempData1, tempData2, emailResult, oFile, newRegDataTemp, enumerationTest, objlogFile, message, error, strLine, objFile, objDict, _
 tempFile2, tempFile1, tempOutput1, tempOutput2, verbose, tempHashData1, tempHashData2, force, argms, i
' --------------------------------------------------

  ' ----------
  ' Company Specific variables.
  ' Change the following variables to match the details of your organization.
  
  ' The " RNscriptName" is the filename of this script.
  RNscriptName = "Registry_Montior.vbs"
  ' The "RNappPath" is the full absolute path for the script directory, with trailing slash.
  RNappPath = "C:\Users\USERNAME\Desktop\Registry_Monitor\"
  ' The "RNlogPath" is the full absolute path for where network-wide logs are stored.
  RNlogPath = "\\SERVER\Logs\"
  ' The "companyName" the the full, unabbreviated name of your organization.
  companyName = "Company Inc."
  ' The "companyAbbr" is the abbreviated name of your organization.
  companyAbbr = "Company"
  ' The "companyDomain" is the domain to use for sending emails. Generated report emails will appear
  ' to have been sent by "COMPUTERNAME@domain.com"
  companyDomain = "Company.com"
  ' The "toEmail" is a valid email address where notifications will be sent.
  toEmail = "IT@Company.com"
  ' Set "emailResult" to TRUE to receive an email when registry modifications are detected. 
  ' Default is TRUE.
  emailResult = TRUE
  ' Set "outputResult" to TRUE to create a lot file when registry modifications are detected. 
  ' Default is TRUE.
  outputResult = TRUE
  ' When "outputResult" is set to TRUE, set "verbose" to TRUE to create a logfile on success or on error (default is error only).
  ' Default is FALSE.
  verbose = FALSE
  ' Set "force" to TRUE to force the script to continue even when it does not have elevated priviledges.
  ' Default is FALSE.
  force = FALSE
  ' ----------

' --------------------------------------------------
'Define frequently used objects.
Set objFSO = CreateObject("Scripting.FileSystemObject") 
Set reg = GetObject("winmgmts://./root/default:StdRegProv")
Set objShell = CreateObject("Wscript.Shell")
Set oShell2 = CreateObject("Shell.Application")
Set objScript = objFSO.GetFile(Wscript.ScriptFullName)
Set argms = WScript.Arguments.Unnamed
Set objDict = CreateObject("Scripting.Dictionary")
'Define constants.
Const HKEY_CLASSES_ROOT = &H80000000
Const HKEY_CURRENT_USER = &H80000001
Const HKEY_LOCAL_MACHINE = &H80000002
Const HKEY_USERS = &H80000003
Const HKEY_CURRENT_CONFIG = &H80000005
'Define variables.
 'Environment/Session related variables.
 strComputerName = objShell.ExpandEnvironmentStrings("%COMPUTERNAME%")
 strUserName = objShell.ExpandEnvironmentStrings("%USERNAME%")
 hiveArray = Array(HKEY_CLASSES_ROOT, HKEY_CURRENT_USER, HKEY_LOCAL_MACHINE, HKEY_USERS, HKEY_CURRENT_CONFIG)
 strKeyPath = ""
 'Date/Time related variables.
 strSafeDate = DatePart("yyyy",Date) & Right("0" & DatePart("m",Date), 2) & Right("0" & DatePart("d",Date), 2)
 strSafeTime = Right("0" & Hour(Now), 2) & Right("0" & Minute(Now), 2) & Right("0" & Second(Now), 2)
 strDateTime = strSafeDate & "-" & strSafeTime
 'File/Directory path related variables.
 scriptPath = objFSO.GetParentFolderName(objScript) 
 logFileName = RNlogPath & strComputerName & "-" & strDateTime & "-Registry_Monitor.txt"
 cachePath = "C:\Users\" & strUserName & "\Registry_Monitor\"
 tempFile1 = cachePath & "Registry_Monitor_Temp1.dat"
 tempFile2 = cachePath & "Registry_Monitor_Temp2.dat"
 regFilePath1 = cachePath & "Registry_Monitor_verifiedKeys.dat"
 regFilePath2 = cachePath & "Registry_Monitor_unverifiedKeys.dat"
 RNmailFile = cachePath & "Registry_Monitor_Warning.mail"
' --------------------------------------------------



' --------------------------------------------------
'Retrieve the specified arguments.
  ' -e  (Email)  =  Set "emailResult" config entry to TRUE (send emails when changes are detected).
  ' -o  (Output)  =  Set "outputResult" config entry to TRUE (create a log file when changes are detected).
  ' -v  (verbose)  =  Set "verbose" config entry to TRUE (create a log file whenever the script is executed).
  ' -f  (Forced)  =  Set "force" config entry to TRUE (bypass elevated priviledge checks, may encounter errors).
Function ParseArgs()
  'Iterate through all supplied arguments.
  For i = 0 to argms.count -1
    'Detect the -e argument.
    If argms.item(i) = "-e" Then
      emailResult = TRUE
    End If
    'Detect the -o argument.
    If argms.item(i) = "-o" Then
      outputResult = TRUE
    End If
    'Detect the -v argument.
    If argms.item(i) = "-v" Then
      verbose = TRUE
    End If
    'Detect the -f argument.
    If argms.item(i) = "-f" Then
      force = TRUE
    End If
  Next 
End Function
' --------------------------------------------------

' --------------------------------------------------
'A function to tell if the script has the required priviledges to run.
'Returns TRUE if the application is elevated.
'Returns FALSE if the application is not elevated.
Function isUserAdmin()
  On Error Resume Next
  CreateObject("WScript.Shell").RegRead("HKEY_USERS\S-1-5-19\Environment\TEMP")
  'Return TRUE if the session is elevated, FALSE if it is not elevated.
  If Err.number = 0 Then 
    isUserAdmin = TRUE
  Else
    isUserAdmin = FALSE
  End If
  Err.Clear
End Function
' --------------------------------------------------

' --------------------------------------------------
'A function to restart the script with admin priviledges if required.
Function restartAsAdmin()
  oShell2.ShellExecute "wscript.exe", Chr(34) & WScript.ScriptFullName & Chr(34), "", "runas", 1
End Function
' --------------------------------------------------

' --------------------------------------------------
'A function to create all required directories before the script can be run & delete any partial files that may already exist.
Function CreateReqdDirs()
  CreateReqdDirs = FALSE
  'Ensure a cachePath exists. Errors at this point probably indicate an intermediary directory does not exist or is not writable.
  If Not objFSO.FolderExists(cachePath) Then
    objFSO.CreateFolder(cachePath)
  End If
  'Ensure a RNlogPath exists. Errors at this point probably indicate an intermediary directory does not exist or is not writable.
  If Not objFSO.FolderExists(RNlogPath) Then
    objFSO.CreateFolder(RNlogPath)
  End If
  'Double check to be sure that required folders were created. 
  If objFSO.FolderExistS(cachePath) And objFSO.FolderExists(RNlogPath) Then
    CreateReqdDirs = TRUE
  End If
  'Delete a pre-existing unverifiedKeys file.
  If objFSO.FileExists(regFilePath2) Then
    objFSO.DeleteFile(regFilePath2)
  End If
End Function
' --------------------------------------------------

' --------------------------------------------------
'A function to create a Warning.mail file. Use to prepare an email before calling sendEmail().
Function createEmail()
  'Check for an existing mail file and delete one if one exists.
  If objFSO.FileExists(RNmailFile) Then
    objFSO.DeleteFile(RNmailFile)
  End If
  'Check for an existing mail file and create one if none exists.
  If Not objFSO.FileExists(RNmailFile) Then
    objFSO.CreateTextFile(RNmailFile)
  End If
  'Set a handle for the "RNmailFile".
  Set oFile = objFSO.CreateTextFile(RNmailFile, True)
  'Write the actual email data to the mail file.
  oFile.Write "To: " & toEmail & vbNewLine & "From: " & strComputerName & "@" & companyDomain & vbNewLine & _
   "Subject: " & companyAbbr & " Registry Monitor Warning!!!" & vbNewLine & _
   "This is an automatic email from the " & companyName & " Network to notify you that a the registry was changed on a domain workstation." & _
   vbNewLine & vbNewLine & "Please verify that the equipment listed below is secure." & vbNewLine & _
   vbNewLine & "USER NAME: " & strUserName & vbNewLine & "WORKSTATION: " & strComputerName & vbNewLine & vbNewLine & _
   "This check was generated by " & strComputerName & "." & vbNewLine & vbNewLine & _
   "Script: """ & RNscriptName & """" 
   'Close the mail file.
  oFile.close
End Function
' --------------------------------------------------

' --------------------------------------------------
'A function for running SendMail to send a prepared Warning.mail email message.
Function sendEmail() 
  objShell.run "c:\Windows\System32\cmd.exe /c sendmail.exe " & RNmailFile, 0, TRUE
End Function
' --------------------------------------------------

' --------------------------------------------------
'A function to create a log file when -l is set.
'Returns "True" if logFilePath exists, "False" on error.
Function CreateRegMonLog(message)
'Make sure the message is not blank.
  If message <> "" Then
    'Set a handle for the "logFileName".
    Set objlogFile = objFSO.CreateTextFile(logFileName, True)
    'Write the "message" to the log file.
    objlogFile.WriteLine(message)
    'Close the log file.
    objlogFile.Close
  End If
  'Check that a lot file was created and return the result.
  If objFSO.FileExists(logFilePath) Then
    error = FALSE
  End If
End Function
' --------------------------------------------------

' --------------------------------------------------
'A function to recursively enumerate registry keys.
'Enumerates through an entire hive & outputs each key to a newline in regFilePath1.
Function EnumerateKeys(hive, key)
  EnumerateKeys = FALSE
  'Enumerate the selected hive/key.
  reg.EnumKey hive, key, arrSubKeys
  'We need to keep re-opening & closing this handle so we don't interfer
  regFileHandle2.WriteLine(hive & "\" & key)
  'If there is a valid subkey to enumerate we call this function again. 
  'This is how we recursively traverse the registry kitting all keys/subkeys in a hive.
  If Not IsNull(arrSubKeys) Then
    For Each subkey In arrSubKeys
      EnumerateKeys hive, key & "\" & subkey
    Next
  End If
End Function 
' --------------------------------------------------



' --------------------------------------------------
'A function to generate a new regFilePath1 if none exists.
'If regFilePath1 does exist, it's contents are loaded & compared with the current registry.
'If discrepencies are found they are reported to the user.
'Once complete this function will replace the existing regFilePath1 with the regFilePath2.
Function VerifyCache()
  newRegData = ""
  'Make sure regFilePath1 exists, & copy regFilePath2 if needed.
  If Not objFSO.FileExists(regFilePath1) Then
    objFSO.CopyFile regFilePath2, regFilePath1
  End If
  'Before we bother iterating through files looking for matches we compare the file hashes to see if they are identical.
  'Output the results of the Certutil command to temporary files.
  objShell.Run "c:\Windows\System32\cmd.exe /c CertUtil -hashfile """ & regFilePath1 & """ SHA256 | find /i /v ""SHA256"" | find /i /v ""certutil"" > """ & tempFile1 & """", 0, TRUE
  Set tempOutput1 = objFSO.OpenTextFile(tempFile1, 1, FALSE, 0)
  objShell.Run "c:\Windows\System32\cmd.exe /c CertUtil -hashfile """ & regFilePath2 & """ SHA256 | find /i /v ""SHA256"" | find /i /v ""certutil"" > """ & tempFile2 & """", 0, TRUE
  Set tempOutput2 = objFSO.OpenTextFile(tempFile2, 1, FALSE, 0)
  'Read the temprary Certutil files & gather the contents (the hashes for the verifiedKeys & unverifiedKeys files).
  If Not tempOutput1.AtEndOfStream Then 
    tempHashData1 = Trim(Replace(Replace(Replace(Replace(tempOutput1.ReadAll(), Chr(10), ""), Chr(13), ""), " ", ""), "  ", ""))
  End If
  If Not tempOutput2.AtEndOfStream Then
    tempHashData2 = Trim(Replace(Replace(Replace(Replace(tempOutput2.ReadAll(), Chr(10), ""), Chr(13), ""), " ", ""), "  ", ""))
  End If
  'Compare the verifiedKeys file hash to the unverifiedKeys file hash & skip the complex comparison if they are identical.
  If tempHashData1 <> tempHashData2 Then
    'Hashes are not identical. Begin complex comparison of cache files.
    Set objFile = objFSO.OpenTextFile(regFilePath1, 1)
    Do Until objFile.AtEndOfStream
      strLine = objFile.ReadLine
      If Not objDict.Exists(strLine) Then
        objDict.Add strLine, ""
      End If         
    Loop
    'Close the "objFile" handle to "regFilePath1".
    objFile.Close
    Set objFile = objFSO.OpenTextFile(regFilePath2, 1)
    Do Until objFile.AtEndOfStream
      strLine = objFile.ReadLine
      If Not objDict.Exists(strLine) Then
        'Build a string of keys that were not found during the search above.
        'Note that we need a separate newRegDataTemp variable to store this data because of the way VBS allocates memory for variables.
        'If you reduce/simplify the declarations below to a single line you will get an "Out of string space" error on most machines.
        newRegDataTemp = newRegData & vbCrLf & strLine
        newRegData = newRegDataTemp
      End If
    Loop
    'Close the "objFile" handle to "regFilePath2".
    objFile.Close
    'When changes are detect; send emails & create a log file if output variables are declared in the configuration section of the script.
    If newRegData <> "" Then
      If outputResult Then
        'Create a log file to warn the user that registry keys have changed.
        CreateRegMonLog("The following registry keys have changed at " & CDate(Now()) & " on machine """ & strComputerName & """ with user """ & strUserName & """: " & vbCrLf & newRegData)
      End If
      If emailResult Then
        'Create an RNmailFile email file using the variables declared in the configuration setting of this script.
        createEmail()
        'Use Sendmail.exe to send the freshly generated RNmailFile warning the user that registry keys have changed.
        sendEmail()
      End If
    End If
  Else
    'If no changes are detected & "verbose" ouputs are enabled, create a log file.
    If verbose And outputResult Then
      'Create the log file.
      CreateRegMonLog("The registry keys on " & strComputerName & " have not been modified on " & CDate(Now()) & ".")
    End If
  End If
  'Close open temp files.
  tempOutput1.Close()
  tempOutput2.Close()
  'Delete any temp files that were created.
  objFSO.DeleteFile(tempFile1)
  objFSO.DeleteFile(tempFile2)
  'Delete the verifiedKeys file.
  objFSO.DeleteFile regFilePath1
  'Copy unverifiedKeys file to verifiedKeys file.
  objFSO.CopyFile regFilePath2, regFilePath1
  'Delete unverifiedKeys file.
  objFSO.DeleteFile(regFilePath2)
End Function
' --------------------------------------------------

' --------------------------------------------------
'The main logic & entry point for the script. Makes use of the functions above.

'Parse the arguments supplied to the script and use them to prepare the operating environment for the session.
'If no arguments are supplied hard-coded configuration entries will be used instead.
parseArgs()

'If "force" variable is set to TRUE; bypass the elevated priviledge check.
If Not force Then
  'Ensure the script has elevated priviledges. 
  If Not isUserAdmin() Then
    'Restart with elevated priviledges if needed.
    restartAsAdmin()
    'Kill the running (non-elevated) script so the elevated one can start & lock needed files.
    WScript.Quit
  End If
End If

'Create directories & clean up cache files.
If CreateReqdDirs() Then
  'Set a handle to the unverifiedKeys file for the EnumerateKeys() function to store registry contents.
  Set regFileHandle2 = objFSO.OpenTextFile(regFilePath2, 8, TRUE, 0)
  'Iterate through each hive & enumerate the keys within.
  For Each hiveItem In hiveArray
    enumerationTest = EnumerateKeys(hiveItem, strKeyPath)
  Next
  'Close the handle to the unverifiedKeys file for the EnumerateKeys() we opened earlier.
  regFileHandle2.Close
  'Check that the unverifiedKeys file was generated by the EnumerateKeys() function.
  If objFSO.FileExists(regFilePath2) Then
    'Compare the enumerated registry keys with the cached version & output the results.
    VerifyCache()
  End If
End If
' --------------------------------------------------
