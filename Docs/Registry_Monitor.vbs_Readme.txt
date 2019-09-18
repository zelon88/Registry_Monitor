T
T
T
NAME: Registry_Monitor

TYPE: VBS Script

PRIMARY LANGUAGE: VBScript
 
AUTHOR: Justin Grimes

ORIGINAL VERSION DATE: 9/5/2019

CURRENT VERSION DATE: 9/18/2019

VERSION: v0.9.8


DESCRIPTION: An application to enumerate registry keys and look for changes which constitute an indicator of compromise.





PURPOSE: To detect malicious registry operations early enough that they do not cause widespread damage to company equipment.




INSTALLATION INSTRUCTIONS: 
1. Install Registry_Monitor into a subdirectory of your Network-wide scripts folder.
2. Open Registry_Monitor.vbs with a text editor and configure the variables at the start of the script to match your environment.
3. Open sendmail.ini with a text editor and configure your email server settings.
4. Run the script automatically with scheduled tasks at regular intervals.




NOTES: 
1. This script MUST be run with administrative rights.
2. If this script is started in regular user mode, it will prompt for administrator elevation.
3. Use absolute UNC paths for network addresses. DO NOT run this from a network drive letter. The restartAsAdmin() function will not work properly.
4. SendMail for Windows is required and included in the "Infrastructure_Heartbeat" folder. The SendMail data files must be included in the same directory as "Infrastructure_Heartbeat.vbs" in order for emails to be sent correctly.