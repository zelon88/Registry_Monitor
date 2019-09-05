'File Name: Registry_Monitor.vbs
'Version: v0.9, 8/29/2019
'Author: Justin Grimes, 8/29/2019

' --------------------------------------------------
Option Explicit
Dim strKeyPath, hive, hiveItem, key, reg, arrSubKeys, subkey, objFSO, regFileHandle1, regFilePath1, regFileHandle2, regFilePath2, hiveArray, objShell, _
 objScript, scriptPath, cachePath,  RNscriptName, RNappPath, RNlogPath, companyName, companyAbbr, companyDomain, toEmail, RNmailFile, logFilePath, Path, _
 reqdDirsExists, newRegData, outputResult, strSafeDate, strSafeTime, strDateTime, logFileName, strComputerName, oShell2
' --------------------------------------------------

  ' ----------
  ' Company Specific variables.
  ' Change the following variables to match the details of your organization.
  
  ' The " RNscriptName" is the filename of this script.
  RNscriptName = "Registry_Montior.vbs"
  ' The "RNappPath" is the full absolute path for the script directory, with trailing slash.
  RNappPath = "\\SERVER\Scripts"
  ' The "RNlogPath" is the full absolute path for where network-wide logs are stored.
  RNlogPath = "\\SERVER\Logs\"
  ' The "companyName" the the full, unabbreviated name of your organization.
  companyName = "Company Inc."
  ' The "companyAbbr" is the abbreviated name of your organization.
  companyAbbr = "Company"
  ' The "companyDomain" is the domain to use for sending emails. Generated report emails will appear
  ' to have been sent by "COMPUTERNAME@domain.com"
  companyDomain = "company.com"
  ' The "toEmail" is a valid email address where notifications will be sent.
  toEmail = "IT@company.com"
  ' The "mailFile" is the full absolute path to the location where a temporary email file will be generated.
  RNmailFile = RNappPath & "Warning.mail"
  ' ----------

' --------------------------------------------------
'Define frequently used objects.
Set objFSO = CreateObject("Scripting.FileSystemObject") 
Set reg = GetObject("winmgmts://./root/default:StdRegProv")
Set objShell = CreateObject("Wscript.Shell")
Set oShell2 = CreateObject("Shell.Application")
Set objScript = objFSO.GetFile(Wscript.ScriptFullName)
'Define constants.
Const HKEY_CLASSES_ROOT = &H80000000
Const HKEY_CURRENT_USER = &H80000001
Const HKEY_LOCAL_MACHINE = &H80000002
Const HKEY_USERS = &H80000003
Const HKEY_CURRENT_CONFIG = &H80000005
'Define variables.
strComputerName = objShell.ExpandEnvironmentStrings("%COMPUTERNAME%")
strUserName = objShell.ExpandEnvironmentStrings("%USERNAME%")
hiveArray = Array(HKEY_CLASSES_ROOT, HKEY_CURRENT_USER, HKEY_LOCAL_MACHINE, HKEY_USERS, HKEY_CURRENT_CONFIG)
strKeyPath = ""
scriptPath = objFSO.GetParentFolderName(objScript) 
cachePath = scriptPath & "\Cache\"
regFilePath1 = cachePath & "verifiedKeys.dat"
regFilePath2 = cachePath & "unverifiedKeys.dat"
strSafeDate = DatePart("yyyy",Date) & Right("0" & DatePart("m",Date), 2) & Right("0" & DatePart("d",Date), 2)
strSafeTime = Right("0" & Hour(Now), 2) & Right("0" & Minute(Now), 2) & Right("0" & Second(Now), 2)
strDateTime = strSafeDate & "-" & strSafeTime
logFileName = RNlogPath & "\" & strComputerName & "-" & strDateTime & "-Registry_Monitor.txt"
' --------------------------------------------------

' --------------------------------------------------
'A function to tell if the script has the required priviledges to run.
'Returns TRUE if the application is elevated.
'Returns FALSE if the application is not elevated.
Function isUserAdmin()
  On Error Resume Next
  CreateObject("WScript.Shell").RegRead("HKEY_USERS\S-1-5-19\Environment\TEMP")
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
'A function to create all required directories before the script can be run.
Function CreateReqdDirs()
  CreateReqdDirs = FALSE
  If Not objFSO.FolderExists(cachePath) Then
    objFSO.CreateFolder(cachePath)
  End If
  If Not objFSO.FolderExists(RNlogPath) Then
    objFSO.CreateFolder(RNlogPath)
  End If
  If objFSO.FolderExistS(cachePath) And objFSO.FolderExists(RNlogPath) Then
    CreateReqdDirs = TRUE
  End If
End Function
' --------------------------------------------------

' --------------------------------------------------
'A function to create a log file when -l is set.
'Returns "True" if logFilePath exists, "False" on error.
Function CreateRegMonLog(message)
  If message <> "" Then
    Set objlogFile = objFSO.CreateTextFile(logFilePath, True)
    objlogFile.WriteLine(message)
    objlogFile.Close
  End If
  If objFSO.FileExists(logFilePath) Then
    error = False
  End If
End Function
' --------------------------------------------------

' --------------------------------------------------
'A function to recursively enumerate registry keys.
'Enumerates the an entire hive and outputs each key to a newline in regFilePath1.
Function EnumerateKeys(hive, key)
  EnumerateKeys = FALSE
  reg.EnumKey hive, key, arrSubKeys
  Set regFileHandle2 = objFSO.OpenTextFile(regFilePath2, 8, TRUE, 0)
  regFileHandle2.WriteLine(hive & "\" & key)
  regFileHandle2.Close
  If Not IsNull(arrSubKeys) Then
    For Each subkey In arrSubKeys
      EnumerateKeys hive, key & "\" & subkey
    Next
  End If
  If objFSO.FileExists(regFilePath2) Then
    EnumerateKeys = TRUE
  End If
End Function 
' --------------------------------------------------

' --------------------------------------------------
'A function to generate a new regFilePath1 if none exists.
'If regFilePath1 does exist, it's contents are loaded and compared with the current registry.
'If discrepencies are found they are reported to the user.
'Once complete this function will replace the existing regFilePath1 with the regFilePath2.
Function VerifyCache()
  newRegData = ""
  If objFSO.FileExists(regFilePath2) Then
    If objFSO.FileExists(regFilePath1) Then
      objFSO.DeleteFile regFilePath1
    End If
    objFSO.CopyFile regFilePath2, regFilePath1
    objFSO.DeleteFile regFilePath2
  End If
  If objFSO.FileExists(regFilePath1) Then
    'Open regFilePath1 and get data as tempData1.
    Set tempHandle1 = objFSO.OpenTextFile(regFilePath1, 1, FALSE)
    tempData1 = tempHandle1.ReadAll()
    'Open regFilePath2 and get data as tempData2.
    Set tempHandle2 = objFSO.OpenTextFile(regFilePath2, 1, FALSE)
    Do Until tempHandle2.AtEndOfStream
      tempData2 = tempHandle2.ReadLine
      If InStr(tempData1, tempData2) > 0 Then
        newRegData = newRegData & vbCrLf & tempData1 & " > " & tempData2
      End If
    Loop
    'Close open files.
    tempHandle1.Close()
    tempHandle2.Close()
    If InStr(tempData2, tempData1) > 0 Then
      messageOutput = "The following registry keys have changed: " & vbCrLf & newRegData
      CreateRegMonLog(messageOutput)
      If outputResult = TRUE Then
        WScript.Echo messageOutput
      End If
    End If
  Else
    objFSO.CopyFile regFilePath2, regFilePath1
  End If
End Function
' --------------------------------------------------

' --------------------------------------------------
'The main logic and entry point for the script. Makes use of the functions above.
If CreateReqdDirs() Then
  For Each hiveItem In hiveArray
    EnumerateKeys hiveItem, strKeyPath
  Next
  VerifyCache()
End If
' --------------------------------------------------