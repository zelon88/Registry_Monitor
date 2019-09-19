NAME: Registry_Monitor

TYPE: VBS Script

PRIMARY LANGUAGE: VBScript
 
AUTHOR: Justin Grimes

ORIGINAL VERSION DATE: 9/5/2019

CURRENT VERSION DATE: 9/19/2019

VERSION: v1.0


DESCRIPTION: An application to enumerate registry keys and look for changes which constitute an indicator of compromise.





PURPOSE: To detect malicious registry operations early enough that they do not cause widespread damage to company equipment.




INSTALLATION INSTRUCTIONS: 
1. Install Registry_Monitor into a subdirectory of your Network-wide scripts folder.
2. Open Registry_Monitor.vbs with a text editor and configure the variables at the start of the script to match your environment.
3. Open sendmail.ini with a text editor and configure your email server settings.
4. Run the script automatically with scheduled tasks at regular intervals.
5. Use the -e argument to force the sending of warning emails.
6. Use the -o argument to force the creation of warning log files.
7. Use the -v argument to force the creation of a log file whenever the script is executed, regardless of detection status.
8. Use the -f argument to force the execution of the script even when the session is not elevated (bypasses elevation checks, may cause errors).




NOTES: 
1. This script MUST be run with administrative rights.
2. If this script is started in regular user mode, it will prompt for administrator elevation.
3. Use absolute UNC paths for network addresses. DO NOT run this from a network drive letter. The restartAsAdmin() function will not work properly.
4. "Fake Sendmail for Windows by Byron Jones" is required and included in the "Registry_Monitor" folder. The SendMail data files must be included in the same directory as "Registry_Monitor.vbs" in order for emails to be sent correctly. 
5. You can download your own copy of "Fake Sendmail for Windows by Byron Jones" by visiting: https://www.glob.com.au/sendmail/.