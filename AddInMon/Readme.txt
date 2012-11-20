What does this program do:  

	This program will monitor the registry and check when an Office Application writes to the StartupItems folder
	within the Resillency registry key of the particular Office Application.  If a value is written to the monitored
	registry folder the application will compare the value to a specified regex pattern to see if it finds a match.  
	If it does it will delete the value from the registry.

	The end result of this process is that when an Office Application hangs/crashes/etc when add-in xyz is active the user
	should no longer receive the "Do you want to disable add-in xyz?" the next time they load the Office Application.

	The program is designed to work with Office 2007/2010 (and possibly 2013)

	

What keys does it monitor:
     Anything written to the Office Apps Resiliency folders:

		HKEY_CURRENT_USER\Software\Microsoft\Office\1?.0\Outlook\Resiliency
		HKEY_CURRENT_USER\Software\Microsoft\Office\1?.0\Excel\Resiliency
		HKEY_CURRENT_USER\Software\Microsoft\Office\1?.0\Word\Resiliency
		HKEY_CURRENT_USER\Software\Microsoft\Office\1?.0\Powerpoint\Resiliency

How do I use this program:
	The program runs as a system tray application in the context of the user who is logged on to the system. 
	When 1st launched it will look for its configuration settings in the following locations:

		1. HKCU\Software\Policies\AddInMon
		2. HKCU\Software\AddInMon
	
	If it finds settings it will load them from the registry and automatically start monitoring the Office applications for
	add-in disables based on the regex pattern.

	If there are no settings the program will just sit in the sytem tray and do nothing until configured.  
	
	You can double click the tray icon to see in realtime what add-ins are being used.  With the program open start 
	using an office application and be amazed as what add-ins are being loaded/unloaded in realtime.  Once you get a feel
	for what applications are running you can create a regular expression to start "protecting" particular addins 
	from being disabled.

What Registry Settings Are there:

	HKCU\Software\AddInMon\Debug	REG_SZ	Values:True or False
		If set this will write out a debug logfile to the users %temp%\AddInMonDebug.txt
	HKCU\Software\AddInMon\RegEx	REG_SZ	Values:Any Regular Expression Pattern
		This value controls what the AddinMon will block from being saved in the StartupItems folder.  Any set of strings
		containing the path or dll name of the add-in you would like to "protect" from being disabled should be listed
		in the regular expression.  
		
		For example we use:		interwoven|interaction|vault


9/1/2012 v 1.1.1.0 [Initial version]

9/25/2012 v 1.1.2.0
	-Fixed a bug that would cause the program to hang when logging off computer
	-Updated HandleEvent Routine to pickup Resiliency keys that may get add/deleted in realtime
	
9/26/2012 v 1.1.3.1
	-Change RegEx pattern match to ignore case
	-Added support for better Office version check if the system does not have Outlook installed
	-Added version label to GUI
	-Added support for Windows XP
	-If a match is found the add-in will show Orange in the real-time list and the monitor is protecting the add-in
    -If it is displaying white then the add-in is not protecting and you need to fix the regex pattern if you want to protect the add-in

	
