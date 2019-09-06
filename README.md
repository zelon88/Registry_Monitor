NAME: Registry_Monitor

TYPE: VBS Script

PRIMARY LANGUAGE: VBScript
 
AUTHOR: Justin Grimes

ORIGINAL VERSION DATE: 9/5/2019

CURRENT VERSION DATE: 9/5/2019

VERSION: v0.9


DESCRIPTION: An application to enumerate registry keys and look for changes which constitute an indicator of compromise.





PURPOSE: To detect malicious registry operations early enough that they do not cause widespread damage to company equipment.




INSTALLATION INSTRUCTIONS: 
1. Install Registry_Monitor into a subdirectory of your Network-wide scripts folder.
2. Open Registry_Monitor.vbs with a text editor and configure the variables at the start of the script to match your environment.
3. Run the script automatically with scheduled tasks at regular intervals.




NOTES: 
1. This script MUST be run with administrative rights.
2. If this script is started in regular user mode, it will prompt for administrator elevation.
3. Use absolute UNC paths for network addresses. DO NOT run this from a network drive letter. The restartAsAdmin() function will not work properly.
4. If using as a startup/logon script it is advised to NOT use a conditional that checks for the prescence of the script prior to running it. Doing so could result in a false negative if ransomware damages Ransomware_Defender before it can be run. Errors produced by such a condition would alert users that something was wrong.