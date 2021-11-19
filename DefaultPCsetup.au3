#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.14.5
 Author:         Jordon Bigelow

 Script Function:
	An automatic New PC Setup Software Installer
	11/18/2021

#ce ----------------------------------------------------------------------------
#RequireAdmin
#include <IE.au3>
#include <File.au3>
AutoItSetOption("TrayIconDebug", 1) ;0-off

#Region Main Body
MsgBox(1, "Auto Installer", "Hello, if anything doesn't function please reach out to Jordon Bigelow.")
CloseFileExplorer()
DisableSecurityPrompt()
UninstallBloatWare()
InstallGoogleChrome()
SetChromeAsDefault()
InstallAdobeReader()
InstallJava()
InstallSymform()
InstallTrendMicro()
InstallDotNetfx()
InstallLaserfiche()
InstallIDScannerDrivers()
InstallLoopback()
InstallSigPlus()
InstallSigWeb()
InstallKaceAgent()
ImportAppAssoc()
InstallOffice()
InstallSilverlight()
InstallPortnox()
InstallExclaimerCloudSignature()
InstallXperience()
SetAdobeAsDefaultPDF()
EnableSecurityPrompt()
MsgBox(1, "Install Complete", "All software has been installed successfully.")
#EndRegion

#Region Functions

Func CloseFileExplorer()
	Sleep(1000)
	WinActivate("Y:\Default New PC Setup\DefaultPC Setup AutoIT Script")
	Sleep(1000)
	WinClose("Y:\Default New PC Setup\DefaultPC Setup AutoIT Script")
EndFunc

Func UninstallBloatWare()
	_FileWriteLog("C:\temp\DefaultPCsetup_log.txt", "Starting UninstallBloatWare function.")
	Run("cmd.exe")
	Sleep(2000)
	WinWait("Administrator: C:\Windows\SYSTEM32\cmd.exe")
	Sleep(1000)
	WinActivate("Administrator: C:\Windows\SYSTEM32\cmd.exe")
	WinWaitActive("Administrator: C:\Windows\SYSTEM32\cmd.exe")
	Send("wmic product where name=""Dell Digital Delivery Services"" call uninstall")
	Send("{ENTER}")
	_FileWriteLog("C:\temp\DefaultPCsetup_log.txt", "Uninstalling Dell Digital Delivery Services.")
	Sleep(2000)
	Send("y")
	Send("{ENTER}")
	Sleep(10000)
	Send("wmic product where name=""Office 16 Click-to-Run Localization Component"" call uninstall")
	Send("{ENTER}")
	Sleep(10000)
	Send("wmic product where name=""Office 16 Click-to-Run Extensibility Component"" call uninstall")
	Send("{ENTER}")
	Sleep(10000)
	Send("wmic product where name=""Office 16 Click-to-Run Licensing Component"" call uninstall")
	Send("{ENTER}")
	sleep(10000)
	WinActivate("Administrator: C:\Windows\SYSTEM32\cmd.exe")
	WinClose("Administrator: C:\Windows\SYSTEM32\cmd.exe")
	_FileWriteLog("C:\temp\DefaultPCsetup_log.txt", "Finished UninstallBloatWare function")
	_FileWriteLog("C:\temp\DefaultPCsetup_log.txt", "")
EndFunc

Func InstallGoogleChrome()
	_FileWriteLog("C:\temp\DefaultPCsetup_log.txt", "Starting InstallGoogleChrome function")
	Local $sProgramName = "Google Chrome"
	_FileWriteLog("C:\temp\DefaultPCsetup_log.txt", "Checking to see if Google Chrome is installed.")
	If CheckIfInstalled($sProgramName) == 0 Then
		_FileWriteLog("C:\temp\DefaultPCsetup_log.txt", "Installing Google Chrome.")
		local $sRunDir = '\\uucu-san\y\Default New PC Setup\'
		; Install File Name
		local $sInstallFile = '1-ChromeSetup.exe'
		ShellExecute('"' & $sRunDir & $sInstallFile & '"')
		WinWait("Installing..., Google Chrome Installer")
		WinWaitActive("Installing..., Google Chrome Installer")
		WinWaitClose("Installing..., Google Chrome Installer", "Installing...")
		WinWait("[CLASS:Chrome_WidgetWin_1]")
		WinActivate("[CLASS:Chrome_WidgetWin_1]")
		WinClose("[CLASS:Chrome_WidgetWin_1]")
		_FileWriteLog("C:\temp\DefaultPCsetup_log.txt", "Finished Installing GoogleChrome.")
	Else
		_FileWriteLog("C:\temp\DefaultPCsetup_log.txt", $sProgramName & " is already installed. Moving on to the next program.")
	EndIf
	_FileWriteLog("C:\temp\DefaultPCsetup_log.txt", "Finished InstallGoogleChrome function.")
	_FileWriteLog("C:\temp\DefaultPCsetup_log.txt", "")
EndFunc   ;==>InstallGoogleChrome

Func SetChromeAsDefault()
	_FileWriteLog("C:\temp\DefaultPCsetup_log.txt", "Starting SetChromeAsDefault function.")
	Run("cmd.exe")
	WinActivate("Administrator: C:\Windows\SYSTEM32\cmd.exe")
	WinWaitActive("Administrator: C:\Windows\SYSTEM32\cmd.exe")
	Send("cd C:\Windows\System32")
	Send("{ENTER}")
	Send("control /name Microsoft.DefaultPrograms /page pageDefaultProgram")
	Send("{ENTER}")
	WinWaitActive("Settings")
	Sleep(1000)
	Send("{TAB 5}")
	Send("{SPACE}")
	Sleep(1000)
	Send("{TAB}")
	Send("{SPACE}")
	Sleep(500)
	Send("{TAB}")
	Send("{SPACE}")
	WinClose("Settings")
	WinActivate("Administrator: C:\Windows\SYSTEM32\cmd.exe")
	WinClose("Administrator: C:\Windows\SYSTEM32\cmd.exe")
	_FileWriteLog("C:\temp\DefaultPCsetup_log.txt", "Finished SetChromeAsDefault function.")
	_FileWriteLog("C:\temp\DefaultPCsetup_log.txt", "")
EndFunc

Func InstallAdobeReader()
	_FileWriteLog("C:\temp\DefaultPCsetup_log.txt", "Starting InstallAdobeReader function.")
	Local $sProgramName = "Adobe Acrobat Reader DC"
	_FileWriteLog("C:\temp\DefaultPCsetup_log.txt", "Checking if Adobe Reader is installed or not.")
	If CheckIfInstalled($sProgramName) == 0 Then
		_FileWriteLog("C:\temp\DefaultPCsetup_log.txt", "Installing Adobe Reader.")
		local $sRunDir = '\\uucu-san\y\Default New PC Setup\'
		; Install File Name
		local $sInstallFile = 'readerdc64_en_xa_crd_install.exe'
		ShellExecute('"' & $sRunDir & $sInstallFile & '"')
		WinWaitActive("Adobe Acrobat Reader DC Installer")
		sleep(150000)
		WinClose("Adobe Acrobat Reader DC Installer")
		;sleep(15000)
		;WinWaitActive("Adobe Acrobat")
		;WinWaitActive("[CLASS:#32770]")
		;ControlClick("[CLASS:#32770]", "", "Button3")
		;WinWaitActive("[CLASS:AcrobatSDIWindow]")
		;WinClose("[CLASS:AcrobatSDIWindow]")
		;Sleep(5000)
		WinWait("Adobe - Download Adobe Acrobat Reader DC - Google Chrome")
		WinActivate("Adobe - Download Adobe Acrobat Reader DC - Google Chrome")
		WinClose("Adobe - Download Adobe Acrobat Reader DC - Google Chrome")
		WinWaitNotActive("Adobe - Download Adobe Acrobat Reader DC - Google Chrome")
		_FileWriteLog("C:\temp\DefaultPCsetup_log.txt", "Adobe Reader is finished installing.")
	Else
		_FileWriteLog("C:\temp\DefaultPCsetup_log.txt", $sProgramName & " is already installed. Moving on to the next program.")
	EndIf
	_FileWriteLog("C:\temp\DefaultPCsetup_log.txt", "Finished InstallAdobeReader function.")
	_FileWriteLog("C:\temp\DefaultPCsetup_log.txt", "")
EndFunc   ;==>InstallAdobeReader

Func SetAdobeAsDefaultPDF(); FIX ME: Get the correct amount of tabs for the .pdf file type.
	_FileWriteLog("C:\temp\DefaultPCsetup_log.txt", "Starting SetAdobeAsDefault function")
	Run("cmd.exe")
	WinWaitActive("Administrator: C:\Windows\SYSTEM32\cmd.exe")
	Sleep(1000)
	Send("cd C:\Windows\System32")
	Send("{ENTER}")
	Sleep(1000)
	Send("control /name Microsoft.DefaultPrograms /page pageDefaultProgram")
	Send("{ENTER}")
	WinWaitActive("Settings")
	Sleep(2000)
	Send("{TAB 7}")
	Send("{SPACE}")
	Sleep(15000)
	Send("{TAB 209}")
	Sleep(2000)
	Send("{TAB 207}")
	Send("{SPACE}")
	Sleep(1000)
	Send("{TAB}")
	Send("{SPACE}")
	Sleep(1000)
	WinClose("Settings")
	WinActivate("Administrator: C:\Windows\SYSTEM32\cmd.exe")
	WinClose("Administrator: C:\Windows\SYSTEM32\cmd.exe")
	WinWaitNotActive("Settings")
	_FileWriteLog("C:\temp\DefaultPCsetup_log.txt", "Finished SetAdobeAsDefault function")
	_FileWriteLog("C:\temp\DefaultPCsetup_log.txt", "")
EndFunc   ;==>SetAdobeAsDefaultPDF

Func InstallJava() ; Installs the latest version of Java
	_FileWriteLog("C:\temp\DefaultPCsetup_log.txt", "Starting InstallJava function")
	Local $sProgramName = "Java 8 Update 311"
	_FileWriteLog("C:\temp\DefaultPCsetup_log.txt", "Checking if Adobe Reader is installed or not.")
	If CheckIfInstalled($sProgramName) == 0 Then
		_FileWriteLog("C:\temp\DefaultPCsetup_log.txt", "Installing Java.")
		Local $sRunDir = '\\uucu-san\y\Default New PC Setup\'
		; Install File Name
		Local $sInstallFile = '3-JavaSetup8u311.exe'
		ShellExecute('"' & $sRunDir & $sInstallFile & '"')
		WinActivate("Java Setup - Welcome")
		WinWaitActive("Java Setup - Welcome")
		; Pushes Tab 5 times to get to the Install button
		Send("{TAB 5}")
		Send("{ENTER}")
		WinWait("Java Setup - Complete")
		WinActivate("Java Setup - Complete")
		WinWaitActive("Java Setup - Complete")
		Send("{ENTER}")
		_FileWriteLog("C:\temp\DefaultPCsetup_log.txt", "Finished Installing Java.")
	Else
		_FileWriteLog("C:\temp\DefaultPCsetup_log.txt", $sProgramName & " is already installed. Moving on to the next program.")
	EndIf
	_FileWriteLog("C:\temp\DefaultPCsetup_log.txt", "Finished InstallJava function.")
	_FileWriteLog("C:\temp\DefaultPCsetup_log.txt", "")
EndFunc   ;==>InstallJava

Func InstallSymform(); FIX ME: Can't use the default registry path to check if symform is installed.
	_FileWriteLog("C:\temp\DefaultPCsetup_log.txt", "Starting InstallSymform function")
	Local $sProgramName = "SymForm"
	_FileWriteLog("C:\temp\DefaultPCsetup_log.txt", "Checking if Symform is installed or not.")
	If CheckIfInstalled($sProgramName) == 0 Then
		_FileWriteLog("C:\temp\DefaultPCsetup_log.txt", "Installing Symform.")
		Local $sRunDir = '\\uucu-san\y\Default New PC Setup\'
		; Install File Name
		Local $sInstallFile = '4-Symform Setup'
		ShellExecute('"' & $sRunDir & $sInstallFile & '"')
		WinActivate("SymForm Setup")
		WinWaitActive("SymForm Setup", "Wizard will install SymForm")
		Send("!n")
		WinWaitActive("SymForm Setup", "Setup will install SymForm")
		Send("!n")
		WinWaitActive("SymForm Setup", "Select the optional components")
		Send("!n")
		WinWaitActive("SymForm Setup", "Please select a program folder.")
		Send("!n")
		WinWaitActive("SymForm Setup", "Installation of SymForm is complete.")
		Send("{ENTER}")
;~ 		_FileWriteLog("C:\temp\DefaultPCsetup_log.txt", "Finished Installing Symform.")
	Else
		_FileWriteLog("C:\temp\DefaultPCsetup_log.txt", $sProgramName & " is already installed. Moving on to the next program.")
	EndIf
	_FileWriteLog("C:\temp\DefaultPCsetup_log.txt", "Finished InstallSymform function.")
	_FileWriteLog("C:\temp\DefaultPCsetup_log.txt", "")
EndFunc   ;==>InstallSymform

Func InstallTrendMicro()
	_FileWriteLog("C:\temp\DefaultPCsetup_log.txt", "Starting InstallTrendMicro function.")
	Local $sProgramName = "Trend Micro Worry-Free Business Security Agent"
	_FileWriteLog("C:\temp\DefaultPCsetup_log.txt", "Checking if TrendMicro is installed or not.")
	If CheckIfInstalled($sProgramName) == 0 Then
		_FileWriteLog("C:\temp\DefaultPCsetup_log.txt", "Installing TrendMicro.")
		Local $sRunDir = '\\uucu-san\y\Default New PC Setup\'
		; Install File Name
		Local $sInstallFile = '5-Trend Micro'
		ShellExecute('"' & $sRunDir & $sInstallFile & '"')
		WinWait("[CLASS:MsiDialogCloseClass]", "dlgbmp")
		WinWaitActive("[CLASS:MsiDialogCloseClass]", "dlgbmp")
		ControlClick("[CLASS:MsiDialogCloseClass]", "dlgbmp", "Button1")
		WinWaitActive("[CLASS:MsiDialogCloseClass]", "Installation Finished")
		ControlClick("[CLASS:MsiDialogCloseClass]", "Installation Finished", "Button1")
	Else
		_FileWriteLog("C:\temp\DefaultPCsetup_log.txt", $sProgramName & " is already installed. Moving on to the next program.")
	EndIf
	_FileWriteLog("C:\temp\DefaultPCsetup_log.txt", "Finished Installing TrendMicro.")
	_FileWriteLog("C:\temp\DefaultPCsetup_log.txt", "")
EndFunc   ;==>InstallTrendMicro

Func InstallDotNetfx()
	_FileWriteLog("C:\temp\DefaultPCsetup_log.txt", "Starting InstallDotNetfx function")
	Local $sRunDir = '\\uucu-san\y\Default New PC Setup\'
	; Install File Name
	Local $sInstallFile = '6-dotnetfx35.exe'
	ShellExecute('"' & $sRunDir & $sInstallFile & '"', "")
	WinWaitActive("Windows Features")
	ControlClick("Windows Features", "", "Button4")
	WinWaitActive("Windows Features", "Installation Results Page")
	ControlClick("Windows Features", "", "Button2")
	_FileWriteLog("C:\temp\DefaultPCsetup_log.txt", "Finished InstallDotNetfx function")
	_FileWriteLog("C:\temp\DefaultPCsetup_log.txt", "")
EndFunc   ;==>InstallDotNetfx

Func InstallLaserfiche()
	_FileWriteLog("C:\temp\DefaultPCsetup_log.txt", "Starting InstallLaserfiche function")
	Local $sProgramName = "Laserfiche 10.4"
	_FileWriteLog("C:\temp\DefaultPCsetup_log.txt", "Checking if Laserfiche is installed or not.")
	If CheckIfInstalled($sProgramName) == 0 Then
		_FileWriteLog("C:\temp\DefaultPCsetup_log.txt", "Starting Laserfiche install.")
		Local $sRunDir = '\\uucu-san\y\Default New PC Setup\'
		; Install File Name
		Local $sInstallFile = '7-SetupLf.exe - Shortcut'
		ShellExecute('"' & $sRunDir & $sInstallFile & '"')
		WinWaitActive("Language Selection")
		ControlClick("Language Selection", "", "Button1")
		WinWaitActive("Laserfiche Windows Client 10.4 Setup", "&Back")
		ControlClick("Laserfiche Windows Client 10.4 Setup","", "Button2")
		WinWaitActive("Laserfiche Windows Client 10.4 Setup", "LASERFICHE END USER LICENSE AGREEMENT")
		ControlClick("Laserfiche Windows Client 10.4 Setup", "", "Button1")
		ControlClick("Laserfiche Windows Client 10.4 Setup", "", "Button4")
		WinWaitActive("Laserfiche Windows Client 10.4 Setup", "Custom")
		ControlClick("Laserfiche Windows Client 10.4 Setup", "Custom", "Button3")
		ControlClick("Laserfiche Windows Client 10.4 Setup", "Custom", "Button8")
		; Navigate the system tree view using the arrows and space bar to uncheck the boxes.
		WinWaitActive("Laserfiche Windows Client 10.4 Setup", "Install directory:")
		Send("{DOWN}")
		Send("{SPACE}")
		Send("{DOWN 3}")
		Send("{SPACE}")
		Send("{DOWN 2}")
		Send("{SPACE}")
		Send("{DOWN}")
		Send("{SPACE}")
		Sleep(10000)
		Send("{DOWN 2}")
		Send("{SPACE}")
		Sleep(1000)
		ControlClick("Laserfiche Windows Client 10.4 Setup", "", "Button11")
		WinWaitActive("Laserfiche Windows Client 10.4 Setup", "Create desktop shortcuts")
		ControlClick("Laserfiche Windows Client 10.4 Setup", "", "Button14")
		WinWaitActive("Laserfiche Windows Client 10.4 Setup", "All prerequisites were successfully installed.")
		ControlClick("Laserfiche Windows Client 10.4 Setup", "", "Button15")
		WinWaitActive("Laserfiche Windows Client 10.4 Setup", "Installation completed successfully.")
		ControlClick("Laserfiche Windows Client 10.4 Setup", "", "Button15")
		WinWaitActive("Laserfiche Windows Client 10.4 Setup", "Finish")
		ControlClick("Laserfiche Windows Client 10.4 Setup", "Finish", "Button18")
		_FileWriteLog("C:\temp\DefaultPCsetup_log.txt", "Finished Installing Laserfiche.")
	Else
		_FileWriteLog("C:\temp\DefaultPCsetup_log.txt", $sProgramName & " is already installed. Moving on to the next program.")
	EndIf
	_FileWriteLog("C:\temp\DefaultPCsetup_log.txt", "Finished InstallLaserfiche function.")
	_FileWriteLog("C:\temp\DefaultPCsetup_log.txt", "")
;~ 	WinClose("Laserfiche 10.4.2")
EndFunc   ;==>InstallLaserfiche

Func InstallIDScannerDrivers()
	_FileWriteLog("C:\temp\DefaultPCsetup_log.txt", "Starting InstallIDScannerDrivers function")
	Local $sProgramName = "PS667"
	_FileWriteLog("C:\temp\DefaultPCsetup_log.txt", "Checking if ID Scanner Drivers are installed.")
	If CheckIfInstalled($sProgramName) == 0 Then
		_FileWriteLog("C:\temp\DefaultPCsetup_log.txt", "Starting ID Scanner Driver install.")
		Local $sRunDir = '\\uucu-mviserver\Tools\Drivers\DocketPort 667\64bit\'
		; Install File Name
		Local $sInstallFile = 'PS667_v1.5.1_setup.exe'
		ShellExecute('"' & $sRunDir & $sInstallFile & '"')
		WinWaitActive("Choose Setup Language")
		ControlClick("Choose Setup Language", "OK", "Button1")
		WinWaitActive("PS667 Setup", "&Next")
		Send("!n")
		WinWaitActive("PS667 Setup", "&Next")
		ControlClick("PS667 Setup", "&Next", "Button1")
		Send("!n")
		WinWaitActive("PS667 Setup", "Destination Folder")
		Send("!n")
		WinWaitActive("PS667 Setup", "&Install")
		Send("!i")
		WinWaitActive("PS667 Setup", "If the New Hardware Wizard")
		Send("!n")
		WinWaitActive("PS667 Setup", "&Finish")
		Send("!f")
		_FileWriteLog("C:\temp\DefaultPCsetup_log.txt", "Finished Installing ID Scanner Drivers.")
	Else
		_FileWriteLog("C:\temp\DefaultPCsetup_log.txt", $sProgramName & " is already installed. Moving on to the next program.")
	EndIf
	_FileWriteLog("C:\temp\DefaultPCsetup_log.txt", "Finished InstallIDScannerDrivers function.")
	_FileWriteLog("C:\temp\DefaultPCsetup_log.txt", "")
EndFunc   ;==>InstallIDScannerDrivers

Func InstallLoopback()
	_FileWriteLog("C:\temp\DefaultPCsetup_log.txt", "Starting InstallIDLoopback function.")
	Local $sRunDir = '\\uucu-mviserver\Tools\Drivers\Loopback\'
	; Install File Name
	Local $sInstallFile = 'RCSPrintPort.reg'
	ShellExecute('"' & $sRunDir & $sInstallFile & '"')
	WinWaitActive("Registry Editor", "&Yes")
	Send("!y")
	WinWaitActive("Registry Editor", "OK")
	ControlClick("Registry Editor", "OK", "Button1")
	Local $sRunDir = '\\uucu-mviserver\Tools\Drivers\Loopback\'
	; Install File Name
	Local $sInstallFile = 'Loopback -Run as Admin.bat'
	ShellExecute('"' & $sRunDir & $sInstallFile & '"', "")
	Sleep(2000)
	_FileWriteLog("C:\temp\DefaultPCsetup_log.txt", "Finished InstallIDLoopback function.")
	_FileWriteLog("C:\temp\DefaultPCsetup_log.txt", "")
EndFunc   ;==>InstallLoopback

Func InstallSigPlus()
	_FileWriteLog("C:\temp\DefaultPCsetup_log.txt", "Starting InstallSigPlus function.")
	Local $sProgramName = "Topaz e-Signatures SigPlus 4.4.0.24"
	_FileWriteLog("C:\temp\DefaultPCsetup_log.txt", "Checking if SigPlus is installed.")
	If CheckIfInstalled($sProgramName) == 0 Then
		_FileWriteLog("C:\temp\DefaultPCsetup_log.txt", "Starting SigPlus install.")
		Local $sRunDir = '\\uucu-mviserver\Tools\Drivers\Sigplus 4.4.0.24\'
		; Install File Name
		Local $sInstallFile = 'sigplus.exe'
		ShellExecute('"' & $sRunDir & $sInstallFile & '"')
		WinWaitActive("Welcome")
		Send("!n")
		WinWaitActive("Read Me File")
		Send("!n")
		WinWaitActive("Choose Destination Location")
		Send("!n")
		WinWaitActive("Determine Tablet Model Group")
		ControlClick("Determine Tablet Model Group", "", "Button3")
		ControlClick("Determine Tablet Model Group", "", "Button7")
		WinWaitActive("Choose the Tablet")
		ControlClick("Choose the Tablet", "", "Button12")
		WinWaitActive("Select the Connection Type")
		ControlClick("Select the Connection Type", "", "Button7")
		WinWaitActive("License Agreement")
		ControlClick("License Agreement", "", "Button1")
		WinWaitActive("Project1")
		Send("{ENTER}")
		WinWaitActive("Demo Ocx.exe")
		ControlClick("Demo Ocx.exe", "", "Button3")
		ControlClick("Demo Ocx.exe", "", "Button4")
		WinWaitActive("Installation Complete")
		Send("!f")
		_FileWriteLog("C:\temp\DefaultPCsetup_log.txt", "Finished Installing Sigplus.")
	Else
		_FileWriteLog("C:\temp\DefaultPCsetup_log.txt", $sProgramName & " is already installed. Moving on to the next program.")
	EndIf
	_FileWriteLog("C:\temp\DefaultPCsetup_log.txt", "Finished InstallSigPlus function.")
	_FileWriteLog("C:\temp\DefaultPCsetup_log.txt", "")
EndFunc   ;==>InstallSigPlus

Func InstallSigWeb()
	_FileWriteLog("C:\temp\DefaultPCsetup_log.txt", "Starting InstallSigWeb function.")
	Local $sProgramName = "Sigweb"
	_FileWriteLog("C:\temp\DefaultPCsetup_log.txt", "Checking if SigWeb is installed.")
	If FileExists("C:\Program Files (x86)\Topaz Systems\SigWeb") == 0 Then
		_FileWriteLog("C:\temp\DefaultPCsetup_log.txt", "Starting SigWeb install.")
		Local $sRunDir = '\\uucu-mviserver\Tools\Drivers\SigWeb 1.7.0.0\sigweb_silent__tpz\'
		Local $sInstallFile = 'SigWeb1.7.0.0.msi'
		ShellExecute('"' & $sRunDir & $sInstallFile & '"')
		WinWait("SigWeb Setup", "&Next >")
		Sleep(1000)
		Send("!n")
		WinWait("SigWeb Setup", "License Agreement and Limited Warranty")
		ControlClick("SigWeb Setup", "IMPORTANT:", "Button4")
		Send("!n")
		WinWait("SigWeb Setup", "Choose the settings for your tablet.")
		Send("{TAB 2}")
		Send("{RIGHT}")
		Send("{TAB 2}")
		Send("{RIGHT 4}")
		Send("{TAB}")
		Send("{RIGHT}")
		Send("!n")
		WinWait("SigWeb Setup", "Topaz Systems, Inc.")
		Send("!n")
		WinWait("SigWeb Setup", "Br&owse...")
		Send("!n")
		WinWait("SigWeb Setup", "Topaz Systems, Inc.")
		Send("!i")
		WinWait("SigWeb Setup", "&Finish")
		Send("!f")
		Sleep(2000)
		WinWait("Topaz SigWeb Certificate Updater")
		WinActivate("Topaz SigWeb Certificate Updater")
		ControlClick("Topaz SigWeb Certificate Updater", "", "Button1")
		Sleep(2000)
		WinWait("Topaz SigWeb Certificate Updater", "OK")
		WinActivate("Topaz SigWeb Certificate Updater", "OK")
		ControlClick("Topaz SigWeb Certificate Updater", "OK", "Button1")
		_FileWriteLog("C:\temp\DefaultPCsetup_log.txt", "Finished Installing SigWeb.")
	Else
		_FileWriteLog("C:\temp\DefaultPCsetup_log.txt", $sProgramName & " is already installed. Moving on to the next program.")
	EndIf
	_FileWriteLog("C:\temp\DefaultPCsetup_log.txt", "Finished InstallSigWeb function.")
	_FileWriteLog("C:\temp\DefaultPCsetup_log.txt", "")
EndFunc   ;==>InstallSigWeb

Func InstallKaceAgent()
	_FileWriteLog("C:\temp\DefaultPCsetup_log.txt", "Starting InstallKaceAgent function.")
	Local $sProgramName = "KACE Agent"
	_FileWriteLog("C:\temp\DefaultPCsetup_log.txt", "Checking if KACE is installed.")
	If CheckIfInstalled($sProgramName) == 0 Then
		_FileWriteLog("C:\temp\DefaultPCsetup_log.txt", "Starting KACE install.")
		Local $sRunDir = '\\uucu-san\y\Default New PC Setup\'
		; Install File Name
		Local $sInstallFile = '9-KACE Agent 11.0.119_UUCU-KBOX.uofucu.com+VRRjpFb9V3evauER4_TB9htapa2udlntpgRn3UWD39bKgWYJzw0jOA.msi - Shortcut'
		ShellExecute('"' & $sRunDir & $sInstallFile & '"')
		WinWait("KACE Agent Setup", "Welcome to the KACE Agent Setup Wizard")
		WinActivate("KACE Agent Setup", "Welcome to the KACE Agent Setup Wizard")
		Sleep(2000)
		Send("!n")
		WinWaitActive("KACE Agent Setup", "Ready to install KACE Agent")
		Send("!i")
		WinWaitActive("KACE Agent Setup", "&Finish")
		Send("!f")
		_FileWriteLog("C:\temp\DefaultPCsetup_log.txt", "Finished Installing Kace Agent.")
	Else
		_FileWriteLog("C:\temp\DefaultPCsetup_log.txt", $sProgramName & " is already installed. Moving on to the next program.")
	EndIf
	_FileWriteLog("C:\temp\DefaultPCsetup_log.txt", "Finished InstallKaceAgent function.")
	_FileWriteLog("C:\temp\DefaultPCsetup_log.txt", "")
EndFunc   ;==>InstallKaceAgent

Func ImportAppAssoc()
	_FileWriteLog("C:\temp\DefaultPCsetup_log.txt", "Starting ImportAppAssoc function")
	Local $sRunDir = '\\uucu-san\y\Default New PC Setup\'
	; Install File Name
	Local $sInstallFile = '11-ImportAppAssoc - Run as Admin.bat'
	ShellExecute('"' & $sRunDir & $sInstallFile & '"')
	Sleep(2000)
	_FileWriteLog("C:\temp\DefaultPCsetup_log.txt", "Finished ImportAppAssoc function")
	_FileWriteLog("C:\temp\DefaultPCsetup_log.txt", "")
EndFunc   ;==>ImportAppAssoc

Func InstallOffice() ; FIX ME: Check back on this function once the whole program is done.
	_FileWriteLog("C:\temp\DefaultPCsetup_log.txt", "Starting InstallOffice function.")
	Local $sProgramName = "Office 16 Click-to-Run Licensing Component"
	_FileWriteLog("C:\temp\DefaultPCsetup_log.txt", "Checking if Office365 is installed.")
	If CheckIfInstalled($sProgramName) == 0 Then
		_FileWriteLog("C:\temp\DefaultPCsetup_log.txt", "Starting Office365 install.")
		Local $sRunDir = '\\uucu-san\y\Default New PC Setup\'
		; Install File Name
		Local $sInstallFile = '12-OfficeSetup.exe'
		ShellExecute('"' & $sRunDir & $sInstallFile & '"')
		Sleep(330000)
		WinWait("[CLASS:Click2RunSetupUIClass]")
		WinActivate("[CLASS:Click2RunSetupUIClass]")
		Send("{ENTER}")
		WinWaitNotActive("‪Microsoft Office‬")
		_FileWriteLog("C:\temp\DefaultPCsetup_log.txt", "Finished Office365 install.")
	Else
		_FileWriteLog("C:\temp\DefaultPCsetup_log.txt", $sProgramName & " is already installed. Moving on to the next program.")
	EndIf
	_FileWriteLog("C:\temp\DefaultPCsetup_log.txt", "Finished InstallOffice function.")
	_FileWriteLog("C:\temp\DefaultPCsetup_log.txt", "")
EndFunc   ;==>InstallOffice

Func InstallSilverlight()
	_FileWriteLog("C:\temp\DefaultPCsetup_log.txt", "Starting InstallSilverlight function")
	Local $sProgramName = "Microsoft Silverlight"
	_FileWriteLog("C:\temp\DefaultPCsetup_log.txt", "Checking if Microsoft Silverlight is installed.")
	If CheckIfInstalled($sProgramName) == 0 Then
		_FileWriteLog("C:\temp\DefaultPCsetup_log.txt", "Starting Microsoft Silverlight install.")
		Local $sRunDir = '\\uucu-san\y\Default New PC Setup\'
		; Install File Name
		Local $sInstallFile = '13-Silverlight_x64 - Shortcut'
		ShellExecute('"' & $sRunDir & $sInstallFile & '"')
		WinWaitActive("Silverlight Setup")
		Send("{TAB 4}")
		Send("{SPACE}")
		Send("{TAB}")
		Send("{SPACE}")
		Send("{TAB 4}")
		Send("{ENTER}")
		WinWaitActive("Silverlight Setup", "Enable Microsoft Update")
		ControlClick("Silverlight Setup", "Enable Microsoft Update", "Button3")
		Sleep(1000)
		ControlClick("Silverlight Setup", "Enable Microsoft Update", "Button4")
		WinWaitActive("Silverlight Setup", "Installation successful")
		ControlClick("Silverlight Setup", "Installation successful", "Button2")
		_FileWriteLog("C:\temp\DefaultPCsetup_log.txt", "Finished Installing Microsoft Silverlight.")
	Else
		_FileWriteLog("C:\temp\DefaultPCsetup_log.txt", $sProgramName & " is already installed. Moving on to the next program.")
	EndIf
	_FileWriteLog("C:\temp\DefaultPCsetup_log.txt", "Finished InstallSilverlight function")
	_FileWriteLog("C:\temp\DefaultPCsetup_log.txt", "")
EndFunc   ;==>InstallSilverlight

Func InstallPortnox()
	_FileWriteLog("C:\temp\DefaultPCsetup_log.txt", "Starting InstallPortnox function")
	Local $sProgramName = "Portnox AgentP"
	_FileWriteLog("C:\temp\DefaultPCsetup_log.txt", "Checking if Portnox is installed.")
	If CheckIfInstalled($sProgramName) == 0 Then
		_FileWriteLog("C:\temp\DefaultPCsetup_log.txt", "Starting Portnox install.")
		Local $sRunDir = '\\uucu-san\y\Default New PC Setup\'
		; Install File Name
		Local $sInstallFile = '14-PortnoxAgentP_1.1.169.exe'
		ShellExecute('"' & $sRunDir & $sInstallFile & '"')
		WinWaitActive("Portnox AgentP")
		; Maps the coords of the window
		; Then moves the mouse to the button relative to the window corner
		Local $aCoord = WinGetPos("Portnox AgentP")
		Local $iX = $aCoord[0] + 445
		Local $iY = $aCoord[1] + 10
		MouseClick("left", $iX, $iY)
		_FileWriteLog("C:\temp\DefaultPCsetup_log.txt", "Finished Installing Portnox.")
	Else
		_FileWriteLog("C:\temp\DefaultPCsetup_log.txt", $sProgramName & " is already installed. Moving on to the next program.")
	EndIf
	_FileWriteLog("C:\temp\DefaultPCsetup_log.txt", "Finished InstallPortnox function")
	_FileWriteLog("C:\temp\DefaultPCsetup_log.txt", "")
EndFunc   ;==>InstallPortnox

Func InstallExclaimerCloudSignature(); FIX ME: Finish this function
	_FileWriteLog("C:\temp\DefaultPCsetup_log.txt", "Starting InstallExclaimerCloudSignature function")
	Local $sRunDir = '\\uucu-san\y\Default New PC Setup\'
	; Install File Name
	Local $sInstallFile = '15-Exclaimer.CloudSignatureUpdateAgent.Install.msi - Shortcut'
	ShellExecute('"' & $sRunDir & $sInstallFile & '"')
	Sleep(15000)
	_FileWriteLog("C:\temp\DefaultPCsetup_log.txt", "Finished InstallExclaimerCloudSignature function")
	_FileWriteLog("C:\temp\DefaultPCsetup_log.txt", "")
EndFunc   ;==>InstallExclaimerCloudSignature

Func InstallXperience()
	_FileWriteLog("C:\temp\DefaultPCsetup_log.txt", "Starting InstallXperience function.")
	Local $sProgramName = "Xperience"
	_FileWriteLog("C:\temp\DefaultPCsetup_log.txt", "Checking if Xperience is installed.")
	If FileExists("C:\Program Files (x86)\JHA") == 0 Then
		_FileWriteLog("C:\temp\DefaultPCsetup_log.txt", "Starting Xperience install.")
		_IECreate("https://jxchange.saltlakecityut.universityfirstfcu.jha-sys.com/Xperience/Infrastructure/Deployment/Production")
		Sleep(5000)
		WinWaitActive("[CLASS:#32770]", "", 5)
		ControlClick("[CLASS:#32770]", "Set up Internet Explorer 11", "Button3")
		Sleep(1000)
		ControlClick("[CLASS:#32770]", "", "Button1")
		; Maps the coords of the window
		; Then moves the mouse to the button relative to the window corner
		Local $aCoord = WinGetPos("Xperience Initializer - Internet Explorer")
		Local $iX = $aCoord[0] + 180
		Local $iY = $aCoord[1] + 290
		MouseClick("left", $iX, $iY)
		WinWaitActive("Application Run - Security Warning")
		Sleep(2000)
		Send("!r")
		Sleep(120000)
		WinWait("Xperience")
		WinWaitActive("Xperience")
		WinClose("Xperience")
		WinClose("Xperience Initializer - Internet Explorer")
		_FileWriteLog("C:\temp\DefaultPCsetup_log.txt", "Finished Installing Xperience.")
	Else
		_FileWriteLog("C:\temp\DefaultPCsetup_log.txt", $sProgramName & " is already installed. Moving on to the next program.")
	EndIf
	_FileWriteLog("C:\temp\DefaultPCsetup_log.txt", "Finished InstallXperience function.")
	_FileWriteLog("C:\temp\DefaultPCsetup_log.txt", "")
EndFunc   ;==>InstallXperience

Func DisableSecurityPrompt()
	_FileWriteLog("C:\temp\DefaultPCsetup_log.txt", "Starting DisableSecurityPrompt function")
	Run("cmd.exe")
	WinWaitActive("Administrator: C:\Windows\SYSTEM32\cmd.exe")
	Send("inetcpl.cpl")
	Send("{ENTER}")
	WinWaitActive("Internet Properties")
	WinClose("Administrator: C:\Windows\SYSTEM32\cmd.exe")
	Send("{TAB 14}")
	Send("{RIGHT}")
	WinWaitActive("Internet Properties", "Custom")
	Send("!c")
	WinWaitActive("Security Settings - Internet Zone")
	Send("{DOWN 133}")
	Send("{SPACE}")
	ControlClick("Security Settings - Internet Zone", "OK", 1)
	WinWaitActive("Warning!")
	Send("!y")
	WinWaitActive("Internet Properties")
	WinClose("Internet Properties")
	_FileWriteLog("C:\temp\DefaultPCsetup_log.txt", "Finished DisableSecurityPrompt function")
	_FileWriteLog("C:\temp\DefaultPCsetup_log.txt", "")
EndFunc

Func EnableSecurityPrompt()
	_FileWriteLog("C:\temp\DefaultPCsetup_log.txt", "Starting EnableSecurityPrompt function")
	Run("cmd.exe")
	WinWaitActive("Administrator: C:\Windows\SYSTEM32\cmd.exe")
	Send("inetcpl.cpl")
	Send("{ENTER}")
	WinWaitActive("Internet Properties")
	WinClose("Administrator: C:\Windows\SYSTEM32\cmd.exe")
	Send("{TAB 14}")
	Send("{RIGHT}")
	WinWaitActive("Internet Properties", "Custom")
	Send("!c")
	WinWaitActive("Security Settings - Internet Zone")
	Send("{DOWN 134}")
	Send("{SPACE}")
	ControlClick("Security Settings - Internet Zone", "OK", 1)
	WinWaitActive("Warning!")
	Send("!y")
	WinWaitActive("Internet Properties")
	WinClose("Internet Properties")
	_FileWriteLog("C:\temp\DefaultPCsetup_log.txt", "Finished EnableSecurityPrompt function")
	_FileWriteLog("C:\temp\DefaultPCsetup_log.txt", "")
EndFunc

Func CheckIfInstalled($sSoftwareName)
	Local $sKey = "HKEY_CLASSES_ROOT\Installer\Products"
	Local $sKey2 = "HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
	Local $sKey3 = "HKLM\Software\Microsoft\Windows\CurrentVersion\Uninstall"

	For $i = 1 To 500
        $AppKey = RegEnumKey($sKey, $i)
		$AppKey2 = RegEnumKey($sKey2, $i)
		$AppKey3 = RegEnumkey($sKey3, $i)
        If @error <> 0 Then ExitLoop
        If StringInStr(RegRead($sKey & "\" & $AppKey, "ProductName"), $sSoftwareName) Then
			Return 1
		ElseIf StringInStr(RegRead($sKey2 & "\" & $AppKey2, "DisplayName"), $sSoftwareName) Then
			Return 1
		ElseIf StringInStr(RegRead($sKey3 & "\" & $AppKey3, "DisplayName"), $sSoftwareName) Then
			Return 1
		EndIf
	Next
	Return 0
EndFunc

Func IsInstalled($progName)
    Local $InstallPrograms = GetList()

    If $InstallPrograms == -1 Then Return -1 ; Cannot get list

    For $program in $InstallPrograms
        if StringInStr(StringLower($program.Caption),StringLower($progName))  Then
            ConsoleWrite($program.Caption & " : " & $progName & @CRLF)
            Return 1;
        EndIf
    Next
    Return -2 ; Didn't find program.
EndFunc

Func GetList()
    $wbemFlagReturnImmediately = 0x10
    $wbemFlagForwardOnly = 0x20
    $colItems = ""
    $strComputer = "localhost"

    $Output=""
    $Output &= "Computer: " & $strComputer  & @CRLF
    $objWMIService = ObjGet("winmgmts:\\" & $strComputer & "\root\CIMV2")
    $colItems = $objWMIService.ExecQuery("SELECT * FROM Win32_Product", "WQL", _
                                              $wbemFlagReturnImmediately + $wbemFlagForwardOnly)

    If IsObj($colItems) Then
        Return $colItems
    EndIf
    Return -1
EndFunc

#EndRegion Functions