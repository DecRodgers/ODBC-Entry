; AutoIT Script to create ODBC entry
#include <MsgBoxConstants.au3>
; Testing shows this app can only write to the registry if run as an admin
#RequireAdmin

;Global Variables

$IniDSNType = IniRead( @ScriptDir & "\ODBC Entry.ini", "ODBC", "DSNType", "No Value Configured")
$IniDSN = IniRead( @ScriptDir & "\ODBC Entry.ini", "ODBC", "DSN", "No Value Configured")
$IniApplication = IniRead( @ScriptDir & "\ODBC Entry.ini", "ODBC", "Application", "No Value Configured")
$IniServer = IniRead( @ScriptDir & "\ODBC Entry.ini", "ODBC", "InstanceName", "No Value Configured")
$IniDB = IniRead( @ScriptDir & "\ODBC Entry.ini", "ODBC", "DatabaseName", "No Value Configured")
$Server32 = RegRead ("HKEY_LOCAL_MACHINE\SOFTWARE\ODBC\ODBC.INI\" & $IniDSN , "Server")
$Database32 = RegRead ("HHKEY_LOCAL_MACHINE\SOFTWARE\ODBC\ODBC.INI\"  & $IniDSN , "Database")
$Server64 = RegRead ("HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\ODBC\ODBC.INI\"  & $IniDSN , "Server")
$Database64 = RegRead ("HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\ODBC\ODBC.INI\"  & $IniDSN , "Database")
$UserDSNDServer = RegRead ("HKEY_CURRENT_USER\Software\ODBC\ODBC.INI\"  & $IniDSN , "Server")
$UserDSNDatabase = RegRead ("HKEY_CURRENT_USER\Software\ODBC\ODBC.INI\"  & $IniDSN , "Database")
$32bitkey = "HKEY_LOCAL_MACHINE\SOFTWARE\ODBC\ODBC.INI\"
$64bitkey = "HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\ODBC\ODBC.INI\"
$UserDSNKey = "HKEY_CURRENT_USER\Software\ODBC\ODBC.INI\"


$arch = @OSArch
;$arch = "X86" - ;Used to test 32 bit on x64 machine
$user = @UserName

IniCheck()
; Checks ini file for blank entries and existance of ini file
Func IniCheck()
   IF $IniDSN = "" OR $IniServer = "" or $IniDB = "" or $IniDSNType = "" or $IniApplication = "" then
	  MsgBox(48, "Configuration error","One or more values not present in ini file. Please check and try again")
	  	  Break(1)
   ElseIf $IniDSN = "No Value Configured" OR $IniServer = "No Value Configured" Or $IniDB = "No Value Configured" or $IniDSNType = "No Value Configured" then
	  MsgBox(48, "Configuration error","Config file not correct format or missing. Please check and try again")
	  	  Break(1)
   Else
	  DSNCheck()
   EndIf
EndFunc

Func DSNCheck()
	If $IniDSNType = "System" Then
		ArchCheck()
	ElseIf $IniDSNType = "User" Then
		ODBCCheckUser()
	Else
		 MsgBox(48, "DSN Type Error","Ini file value '" & $IniDSNType & "' not valid. Please check and try again")
			Break(1)
	EndIf
EndFunc


;Checks OS Architecture
Func ArchCheck()
   Select
   ; X86
	  Case $arch = "X86"
		 ODBCCheck32()
   ; X64
	  Case $arch = "X64"
		 ODBCCheck64()
   ; Unverified
	  Case Else
		 Uknown()
   EndSelect
EndFunc

;Checks for existance of ODBC entry with the same entry as ini file and goes to an optional overwrite section if there is

Func ODBCCheck32()
If $Server32 <> "" Then
	  Overwrite32()
   Else
	  ThirtyTwobit()
EndIf
EndFunc

Func ODBCCheck64()
If $Server64 <> "" Then
	  Overwrite64()
   Else
	  SixtyFourbit()
EndIf
EndFunc

Func ODBCCheckUser()
If $UserDSNDServer <> "" Then
		OverWriteUser()
	Else
		UserDsnInstall()
	EndIf
EndFunc

;Gives user option to overwrite ODBC entry if it currently exists

Func Overwrite32()
$ThirtyTwoOverwrite = MsgBox(4, "System DSN Entry '" & $IniDSN & "' exists", "You currently have an entry for '" & $Server32 & "' and database '" & $Database32 & "'. Would you like to overwrite this?")
   Select
   Case $ThirtyTwoOverwrite = $IDYES
	  ThirtyTwobit()
   Case $ThirtyTwoOverwrite = $IDNO
	  Cancel()
   EndSelect
EndFunc

Func Overwrite64()
$SixtyFourOverwrite = MsgBox(4, "System DSN Entry '" & $IniDSN & "' exists", "You currently have an entry for '" & $Server64 & "' and database '" & $Database64 & "'. Would you like to overwrite this?")
   Select
   Case $SixtyFourOverwrite = $IDYES
	  SixtyFourbit()
   Case $SixtyFourOverwrite = $IDNO
	  Cancel()
   EndSelect
EndFunc

Func OverWriteUser()
$UserDSN = MsgBox(4, "User DSN Entry '" & $IniDSN & "' exists", "You currently have an entry for '" & $UserDSNDServer & "' and database '" & $UserDSNDatabase & "'. Would you like to overwrite this?")
   Select
   Case $UserDSN = $IDYES
	  SixtyFourbit()
   Case $UserDSN = $IDNO
	  Cancel()
   EndSelect
EndFunc


;Sections to write entry into registry

Func ThirtyTwobit()
$Install32 = MsgBox(4, "32 Bit detected", "Would you like to install an ODBC entry of '"& $IniDSN & "' to '" & $IniServer & "' and database '" & $IniDB & "'?")
   Select
	  Case $Install32 = $IDYES
		 RegWrite ( $32bitkey & "ODBC Data Sources\", $IniDSN , "REG_SZ", "SQL Server")
		 RegWrite ( $32bitkey & $IniDSN , "Driver", "REG_SZ", "C:\Windows\system32\SQLSRV32.dll")
		 RegWrite ( $32bitkey & $IniDSN , "Server", "REG_SZ", $IniServer)
		 RegWrite ( $32bitkey & $IniDSN , "Database", "REG_SZ", $IniDB)
		 RegWrite ( $32bitkey & $IniDSN , "Description", "REG_SZ", $IniApplication & " Connection")
		 RegWrite ( $32bitkey & $IniDSN , "LastUser", "REG_SZ", $User)
		 RegWrite ( $32bitkey & $IniDSN , "Trusted_Connection", "REG_SZ", "Yes")
			Installed()
	  Case $Install32 = $IDNO
		 Cancel()
   EndSelect
EndFunc

Func SixtyFourbit()
$Install64 = MsgBox(4, "64 Bit detected", IniApplication & " needs to use a 32 bit ODBC connection. Would you like to install a 32 bit ODBC entry of '"& $IniDSN & "' to '" & $IniServer & "' and database '" & $IniDB & "'?")
   Select
	  Case $Install64 = $IDYES
		 RegWrite ( $64bitkey & "ODBC Data Sources\", $IniDSN , "REG_SZ", "SQL Server")
		 RegWrite ( $64bitkey & $IniDSN , "Driver", "REG_SZ", "C:\Windows\system32\SQLSRV32.dll")
		 RegWrite ( $64bitkey & $IniDSN , "Server", "REG_SZ", $IniServer)
		 RegWrite ( $64bitkey & $IniDSN , "Database", "REG_SZ", $IniDB)
		 RegWrite ( $64bitkey & $IniDSN , "Description", "REG_SZ", $IniApplication & " Connection")
		 RegWrite ( $64bitkey & $IniDSN , "LastUser", "REG_SZ", $User)
		 RegWrite ( $64bitkey & $IniDSN , "Trusted_Connection", "REG_SZ", "Yes")
			Installed()
	  Case $Install64 = $IDNO
		 Cancel()
   EndSelect
EndFunc

Func UserDsnInstall()
$InstallUser  = MsgBox(4, "Installation Confirmation", "Would you like to install a User ODBC entry of '"& $IniDSN & "' to '" & $IniServer & "' and database '" & $IniDB & "'?")
   Select
	  Case $InstallUser = $IDYES
		 RegWrite ( $UserDSNKey & "ODBC Data Sources\", $IniDSN , "REG_SZ", "SQL Server")
		 RegWrite ( $UserDSNKey & $IniDSN , "Driver", "REG_SZ", "C:\Windows\system32\SQLSRV32.dll")
		 RegWrite ( $UserDSNKey & $IniDSN , "Server", "REG_SZ", $IniServer)
		 RegWrite ( $UserDSNKey & $IniDSN , "Database", "REG_SZ", $IniDB)
		 RegWrite ( $UserDSNKey & $IniDSN , "Description", "REG_SZ", $IniApplication & " Connection")
		 RegWrite ( $UserDSNKey & $IniDSN , "LastUser", "REG_SZ", $User)
		 RegWrite ( $UserDSNKey & $IniDSN , "Trusted_Connection", "REG_SZ", "Yes")
			Installed()
	  Case $InstallUser = $IDNO
		 Cancel()
   EndSelect
EndFunc

Func Installed()
   MsgBox(48, "Entry installed","ODBC information has now been installed.")
EndFunc

Func Cancel()
   MsgBox(48, "Canceled by user","ODBC entry not installed.")
EndFunc

; if it detects an architecture that is uknown,program stops, for System, DSNs only

Func Uknown()
MsgBox(16, "Error", "Unspecified architecture, program will now close.")
	  Break(1)
EndFunc