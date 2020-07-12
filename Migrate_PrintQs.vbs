' ***************************************************************************************************
' ***************************************************************************************************
'
'  Author       : Cary GARVIN
'  Contact      : cary(at)garvin.tech
'  LinkedIn     : https://www.linkedin.com/in/cary-garvin-99909582
'  GitHub       : https://github.com/carygarvin/
'
'
'  Script Name  : Migrate_PrintQs.vbs
'  Version      : 1.0
'  Release date : 07/02/2015 (CET)
'  History      : The present script has been used by large organizations to successfully migrate tens of thousands of network printers from old to new Print Servers. A lot of safeguards have been buil into this Script.
'  Purpose      : The present script is to be used in the scope of a Print Server migration whereby Network Print Queues are migrated from one old Print Server (to be decommissioned) to a new Print server.
'                 The script will take care of remapping all of user Print Queues based on information contained in a mappings or correspondence file with each line in the format '\\OldPrintServer\OldPrintQueueName,\\NewPrintServer\NewPrintQueueName'.
'                 On top of that the script has the ability to remove (Remove) some printers altogether or Add (Affix) one or more printers deemed compulsory.
'
'
'	Script to migrate user printers from one Print Server to another based on a correspondence/mapping file holding Old Print Queue to New Print Queue mappings.
'	The present Script is best invoked during the Login Script through Active Directory Group membership or interactively by specifying any specific mapping file to use as a parameter
'   The present Script uses for maximum reliability several methods in order to identify user printers.
'   The present Script has many features as follows:
'                  * Migrate user PrintQueues at logon or interactively based on information contained in the specified mappings file
'                  * Add one or more compulsory printers which all users not to have access to.
'                  * Unequivocally remove obsolete printers.
'
'   As stated, the present script can be invoked either from a Logon Script in which the current user can have its group membership tested and if validated, call the script. By default, the mappings file that will be used will match the user's OU. This allows to have mappings file specific to each departmernt/bu in case of large organizations.
'   Here's an example of how it can be called from within a "parent" VBScript Logon Script provided the Group's Distinguished Name is stored in strGroupDN and a binding has been made to the user object through objUser:
'                          Set objGroup = GetObject("LDAP://" & strGroupDN)
'                          If objGroup.IsMember("LDAP://" & objUser.UserName) Then
'                              objShell.Run "Migrate_PrintQs.vbs"
'                          EndIf 
'
'   Alternatively, the script can be run from a Command Line (cscript Migrate-PrintQs.vbs). 
'   The script's Remove or Affix feature can be invoked either through Command Line switches when invoking the script (mostly used in interactive cases) or through specific fomratting of mappings with the mappings file (mostly used via a Logon Script).
'
'
'   Command Line switches: 
'                           FileName.csv
'                           /Affix:
'                           /Remove:
'                           /RemoveAllPrinters
'                           /CheckGroupMembership
'                           /CheckGroupMembership:<CustomGroupName>
'
'   Command Line Examples:
'                           Migrate_PrintQs.vbs PrintMigTable.csv                                   ==>      [This will migrate current Print Queues based on the information inside specified 'PrintMigTable.csv' file. This file is to be posted on the Network Share specified in the 'PrintQMappingsRepo' variable]
'                           Migrate_PrintQs.vbs /Affix:\\ContosoNewPrtSrv\NewPrintQueueName			==>      [This will add a mapping to '\\ContosoNewPrtSrv\NewPrintQueueName' if none already exists]
'                           Migrate_PrintQs.vbs /Remove:\\ContosoOldPrtSrv\OldPrintQueueName		==>      [This will remove any mapping to '\\ContosoOldPrtSrv\OldPrintQueueName' if any exists]
'                           Migrate_PrintQs.vbs /RemoveAllPrinters                                  ==>      [This will remove all of user's printers]
'                           Migrate_PrintQs.vbs /CheckGroupMembership                               ==>      [This will tell the script to act as if it is run within the Logon Script, meaning that the Mappings table to use is the default computed one for the user's devised Department.]
'                           Migrate_PrintQs.vbs /CheckGroupMembership:PrintMigUsers                 ==>      [Same as above but for special cases where the user does not comply to the Department OU = Group prefix = Mappings CSV file prefix paradigm. The migration will take place based on the Mappings table from the user's Department OU]
'
'   Migration action Examples via mappings file (assuming Print Server migration from 'ContosoOldPrtSrv1' to 'ContosoNewPrtSrv1'):
'                           \\ContosoOldPrtSrv1\OldPrtQ1,\\ContosoNewPrtSrv1\NewPrtQ1               ==>      Print Queue '\\ContosoOldPrtSrv1\OldPrtQ1' will be replaced by '\\ContosoNewPrtSrv1\NewPrtQ1'
'                           \\ContosoOldPrtSrv1\OldPrtQ2,\\ContosoNewPrtSrv1\NewPrtQ2               ==>      Print Queue '\\ContosoOldPrtSrv1\OldPrtQ2' will be replaced by '\\ContosoNewPrtSrv1\NewPrtQ2'
'                           \\ContosoOldPrtSrv1\OldPrtQ3,\\ContosoNewPrtSrv1\NewPrtQ3               ==>      Print Queue '\\ContosoOldPrtSrv1\OldPrtQ3' will be replaced by '\\ContosoNewPrtSrv1\NewPrtQ3'
'                           \\ContosoOldPrtSrv1\OldPrtQ4,                                           ==>      Print Queue '\\ContosoOldPrtSrv1\OldPrtQ4' will be univocally removed if found
'                           \\ContosoOldPrtSrv1\OldPrtQ5,DELETE                                     ==>      Print Queue '\\ContosoOldPrtSrv1\OldPrtQ5' will be univocally removed if found
'                           \\ContosoNewPrtSrv1\NewGrpPrtQ,INSTALL                                  ==>      Print Queue '\\ContosoNewPrtSrv1\NewGrpPrtQ' will be univocally added if not found
'
'
'  There are 5 configurable variables (see lines 149 to 153) which need to be set by IT Administrator prior to using the present Script:
'  Variable 'DeptsOU' contains the parent node OU in the form "OU=xyz" where all departments are residing.
'  Variable 'PrintServersOU' contains the OU where the Print Servers involved in the migration are located. Specifying this allows for fatser LDAP searches.
'  Variable 'PrintMigGroupsOU' contains the OU where the different printer migrations Groups are residing. Specifying this allows for fatser LDAP searches.
'            Printer Migrations Groups for each BU/Department/OU are expected to match the pattern "<SubOUNameInDeptsOUVar>-PrinterMigration".
'            So for instance for HR, the script expects an 'HR' OU inside 'DeptsOU' above and the AD Group containg HR users which can migrate to be named "HR-PrinterMigration" and reside in AD inside the OU specified through this "PrintMigGroupsOU" variable.
'  Variable 'PrintQMappingsRepo' contains the UNC location of where the mapping file(s) reside. (Ensure NTFS Security and Share permissions are set for 'Everyone' to READ).
'            Printer Mappings files to be posted here are expected to match the string pattern "<SubOUNameIn{DeptsOU}Var>-PrintQMig.csv". 
'            So for instance again for HR, the script expects an 'HR' OU inside 'DeptsOU' above and the file containing the mappings for HR must be called "HR-PrintQMig.csv" and obviously must be present on the Network Share specified in this 'PrintQMappingsRepo'.
'  Variable 'PrintQMigrationLogs' contains the UNC location of where the user Print Queue migrations logs are to be created. (Ensure NTFS Security and Share permissions are set for 'Everyone' to WRITE).
'
'
'  Note: The bahaviour for user's Default Printer is that if whatever reason no new Default Printer can be set in its place, it will never be removed.
'
'  ****************************************************************************************************
'  ****************************************************************************************************


Option Explicit


'Initializations and Constructions
Const ForReading = 1, ForWriting = 2, ForAppending = 8
Const HKEY_CURRENT_USER = &H80000001
Const HKEY_LOCAL_MACHINE = &H80000002
Const PRINTERS_AND_FAXES = &H4&
Const DEFAULT_PRINTER = 4

Public TSSession : TSSession = vbNullString
Public OSVersion
Public PrinterToAffix : PrinterToAffix = vbNullString
Public PrinterToRemove : PrinterToRemove = vbNullString
Public CustomGroup : CustomGroup = vbNullString
Public DefaultPrinterName : DefaultPrinterName = vbNullString
Public netPrinters
Public netPrintersCount
Public MappingFileError : MappingFileError = vbNullString
Public MappingFileWarning : MappingFileWarning = vbNullString
Public PrintMigGroupError : PrintMigGroupError = vbNullString
Public PrintMigGroupValidationLogLineError : PrintMigGroupValidationLogLineError = vbNullString
Public SingleTaskStatus : SingleTaskStatus = vbNullString
Public MappingTable, AllMappings
Public EnrolledPrinters : EnrolledPrinters = 0
Public RejectedPrinters : RejectedPrinters = 0
Public MigratedPrinters : MigratedPrinters = 0
Public AffixedPrinters : AffixedPrinters = 0
Public RemovedPrinters : RemovedPrinters = 0
Public UnremovablePrinters : UnremovablePrinters = 0
Public PrintersToAffix : PrintersToAffix = vbNullString
Public InstalledPrintersList : InstalledPrintersList = vbNullString
Public AddedPrintersList : AddedPrintersList = vbNullString
Public boolFinished

Dim TimeTrack(2)
Dim objShell : Set objShell = WScript.CreateObject("WScript.Shell")
Dim objShellApp : Set objShellApp = CreateObject("Shell.Application")
Dim objFSO : Set objFSO = CreateObject("Scripting.FileSystemObject")
Dim objWMIService : Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
Dim objNetwork : Set objNetwork = CreateObject("WScript.Network")
Dim objConnection : Set objConnection = CreateObject("ADODB.Connection")  
Dim objCommand : Set objCommand = CreateObject("ADODB.Command")  
Dim objRootDSE, objRecordSet
Dim strUserName, strComputerName, strStationType
Dim strDomainDN , strDomain, strUserDN, strUserOU, strUserCN, strUserDept
Dim UserGroupList : UserGroupList = vbNullString
Dim arrUserGroupList
Dim strCurrentPrinterShortUNC, arrOldToNewPrinter, strOldPrinterUNC, strNewPrinterUNC
Dim PrintMigUnderTS : PrintMigUnderTS = True
Dim GreenLight : GreenLight = False
Dim PrintMigrateUser : PrintMigrateUser = False
Dim SinglePrinterTask : SinglePrinterTask = False
Dim WipeAllPrintersTask : WipeAllPrintersTask = False
Dim Argument
Dim InteractiveMode : InteractiveMode = False
Dim ConditionalMode : ConditionalMode = False
Dim colMemberOf
Dim MappingFile, MappingFileName, DestLogDir, UserLogTimeStamp, UserLogFile, objMappingFile
Dim objOutput
Dim i, ii, j
Dim InconsistentLines : InconsistentLines = vbNullString
Dim InconsistentLineCount : InconsistentLineCount = 0
Dim InconsistencyTracker : InconsistencyTracker = vbNullString
Dim CurrentFieldSeparator, InvalidSeparatorLineCount
Dim AllFieldSeparators : AllFieldSeparators = vbNullString

' Configurable parameters by IT Administrator as referred in the script's instructions. Modify these 5 varaiable values to match your AD and network configuration.
Dim DeptsOU : DeptsOU = ",OU=Departments,"                                                                  ' The OU containing all of your Organization's Business Units (aka OUs/Deparments)
Dim PrintServersOU : PrintServersOU = "<LDAP://OU=Print Servers,OU=Servers,OU=Computers,OU=IT,"             ' Don't specify the trailing Domain portion as it is discovered dynamically
Dim PrintMigGroupsOU : PrintMigGroupsOU = "<LDAP://OU=Groups,OU=IT,OU=PrintMigDepartmentGroups,"            ' Don't specify the trailing Domain portion as it is discovered dynamically
Dim PrintQMappingsRepo : PrintQMappingsRepo = "\\ContosoSrv1\IT\PrintQMig\DeptMappings\"                    ' The Print Q mappings file is expected to be called '<DepartmentName>-PrintQMig.csv' where '<DeparmentName>' is extracted from the OU container in the users's Distinguished Name. (See line ???? - adjust to suit your needs).
Dim PrintQMigrationLogs : PrintQMigrationLogs = "\\ContosoSrv1\IT\PrintQMig\UserLogs\"                      ' The network location where user individual migration logs are stored. These logs allow IT Administrators to monitor the progress of the Print Queue migrations

'End of variable and object initializations. Script main body below
'####################################################################################################


TimeTrack(0) = TimeTracker


'ADODB objects for LDAP queries (to check if printer is published in AD, query user's Department, etc...)
objConnection.Provider = "ADsDSOObject"  
objConnection.Open "Active Directory Provider"  
objCommand.ActiveConnection = objConnection


strUserName = UCase(objShell.ExpandEnvironmentStrings("%USERNAME%"))
strComputerName = UCase(objShell.ExpandEnvironmentStrings("%COMPUTERNAME%"))


Set objRootDSE = GetObject("LDAP://RootDSE")
strDomainDN = UCase(objRootDSE.Get("defaultNamingContext"))
strDomain = Replace(Mid(strDomainDN,4),",DC=",CHR(46))
Set objRootDSE = nothing


'Get from logged on user his Active Directory Distinguished Name and Common/Display Name 
objCommand.CommandText = "<LDAP://" & strDomainDN & ">;(&(objectCategory=person)(objectClass=user)(sAMAccountName=" & strUserName & "));distinguishedName,cn;subtree" 
Set objRecordSet = objCommand.Execute
strUserDN = UCase(objRecordSet.Fields("distinguishedName").Value)
strUserCN = objRecordSet.Fields("cn").Value


'Determine user's OU and Department name. If the user is in the "Departments" OU () tree extract the Department acronym, else default to the OTHER mechanism 
If InStr(strUserDN,DeptsOU) Then
	strUserOU = Left(strUserDN,Len(strUserDN)-Len(DeptsOU & strDomainDN))
	strUserDept = Mid(strUserOU,InStrRev(strUserOU,"=",-1)+1)
Else
	strUserDept = "OTHER"
End If


'Mapping table selection and migration log initialization. A specific mapping file can be used by passing it as a script parameter. If none is specified, use the mapping file corresponding to the user's Department as identified by the Department OU and subsequently stored in variable 'strUserDept'.
If WScript.Arguments.Count = 0 Then
	'Department dependant Printer migration mapping table if no argument is passed on when invoking present script. 
	SetLogDir "AutoSelection", "PrintMig", strUserDept
	MappingFile = PrintQMappingsRepo & strUserDept & "-PrintQMig.csv"
	MappingFileName = strUserDept & "-PrintQMig.csv"
	
	'With no argument passed when calling the script, the script is assumed called from the Logon Script and/or the mapping file to use is that of the user's Department
	If DeptPrintMigGroupExists(strUserDept) Then
		GreenLight = True
		PrintMigrateUser = True
	End If
Else
	For Each Argument In WScript.Arguments
		If InStr(lCase(Argument),"/help") or InStr(lCase(Argument),"/?")Then
			WScript.Echo "The Script syntax is as follows when using arguments where brackets indicate"
			WScript.Echo "free values (don't use brackets in the actual command line):"
			WScript.Echo
			WScript.Echo "cscript Migrate_PrintQs.vbs [AnyFileName.csv]"
			WScript.Echo vbTab & "to run the script against specific mapping file AnyFileName.csv"
			WScript.Echo
			WScript.Echo "cscript Migrate_PrintQs.vbs /CheckGroupMembership"
			WScript.Echo vbTab & "to only run the script against default mapping file " & strUserDept & "-PrintQMig.csv"
			WScript.Echo vbTab & "for users members of AD group " & strUserDept & "-PrinterMigration"
			WScript.Echo
			WScript.Echo "cscript Migrate_PrintQs.vbs /CheckGroupMembership:[GroupName]"
			WScript.Echo vbTab & "to only run the script against mapping file " & strUserDept & "-PrintQMig.csv for users"
			WScript.Echo vbTab & "members of AD group [GroupName]"
			WScript.Echo
			WScript.Echo "cscript Migrate_PrintQs.vbs /Affix:[\\PrinterServer\PrintQueue]"
			WScript.Echo vbTab & "to add printer \\PrinterServer\PrintQueue for all users without it"
			WScript.Echo
			WScript.Echo "cscript Migrate_PrintQs.vbs /Remove:[\\PrinterServer\PrintQueue]"
			WScript.Echo vbTab & "to remove printer \\PrinterServer\PrintQueue for all users having"
			WScript.Echo vbTab & "printer [\\PrinterServer\PrintQueue] installed"
			WScript.Echo
			WScript.Echo "cscript Migrate_PrintQs.vbs /RemoveAllPrinters"
			WScript.Echo vbTab & "to remove ALL installed network printers for all users"
			WScript.Echo
			WScript.Echo "Additionnally, the /Affix:, the /Remove: or the /RemoveAllPrinters switches can jointly"
			WScript.Echo "be used with either of the two possible /CheckGroupMembership variants"
		End If
		
		If InStr(lCase(Argument),".csv") Then
			GreenLight = True
			PrintMigGroupError = "NoGroupCheck"
			If lCase(Argument) = "all-PrintQMig.csv" Then
				PrintMigrateUser = True
				MappingFile = PrintQMappingsRepo & "ALL-PrintQMig.csv"
				MappingFileName = "ALL-PrintQMig.csv"
				DestLogDir = PrintQMigrationLogs & "PrintMig-ALL"
			Else
				InteractiveMode = True
				If lCase(Right(Argument,4)) = ".csv" Then
					'Used for instance for organization wide Printer migration with global mapping table passed on as argument when invoking current script.
					PrintMigrateUser = True
					MappingFile = PrintQMappingsRepo & Argument
					MappingFileName = Argument
				Else
					MappingFileError = "Printer mapping file " & Argument & " passed on as script parameter is not a valid mapping file name. No printer migration performed! Please check file name!"
					DestLogDir = PrintQMigrationLogs & "Interactive\Bad"
				End If
			End If
		End If
	
		If InStr(lCase(Argument),"/checkgroupmembership") Then
			ConditionalMode = True
			If InStr(Argument,CHR(58)) = 0 Then
				If DeptPrintMigGroupExists(strUserDept) Then
					If GroupMember (strUserDept & "-PrinterMigration") Then
						'Flag to Migrate if current user is a member of the default '<DeptName>-PrinterMigration' Group (for Logon Script migrations, see comments above, this check is done in the Logon Script as it invokes 'Migrate_PrintQs.vbs' only if the user is found member of the Department Printer Migration Group for controlled phased migrations within each department)
						GreenLight = True
					End If
				End If
			Else
				CustomGroup = Mid(Argument,Instr(Argument,CHR(58))+1)
				If GroupMember (CustomGroup) Then
					'Flag to Migrate if current user is a member of the custom group specified on the command line (Usually for non Logon Script Migrations)
					GreenLight = True
				End If
			End If
			If IsArray(arrUserGroupList) Then
				Erase arrUserGroupList
			End If
			'Setting of Department dependant Printer migration mapping table and log file repository
			If SinglePrinterTask = False and WipeAllPrintersTask = False Then
				MappingFile = PrintQMappingsRepo & strUserDept & "-PrintQMig.csv"
				MappingFileName = strUserDept & "-PrintQMig.csv"
			End If
		End If
		
		If InStr(lCase(Argument),"/affix:") Then
			MappingFile = vbNullString
			MappingFileName = vbNullString
			SinglePrinterTask = True
			PrinterToAffix = UCase(Mid(Argument,Instr(Argument,CHR(58))+1))
			If WScript.Arguments.Count = 1 Then
				GreenLight = True
			End If
		End If
		
		If InStr(lCase(Argument),"/remove:") Then
			MappingFile = vbNullString
			MappingFileName = vbNullString
			SinglePrinterTask = True
			PrinterToRemove = UCase(Mid(Argument,Instr(Argument,CHR(58))+1))
			If WScript.Arguments.Count = 1 Then
				GreenLight = True
			End If
		End If
		
		If InStr(lCase(Argument),"/removeallprinters") Then
			MappingFile = vbNullString
			MappingFileName = vbNullString
			WipeAllPrintersTask = True
			If WScript.Arguments.Count = 1 Then
				GreenLight = True
			End If
		End If
	Next
End If


If GreenLight Then
	If WScript.Arguments.Count > 0 Then
		If SinglePrinterTask or WipeAllPrintersTask Then
			If SinglePrinterTask Then
				If ConditionalMode Then
					SetLogDir "Conditional", "SingleTask", strUserDept
				Else
					SetLogDir "AutoSelection", "SingleTask", strUserDept
				End If
			Else
				If ConditionalMode Then
					SetLogDir "Conditional", "WipeTask", strUserDept
				Else
					SetLogDir "AutoSelection", "WipeTask", strUserDept
				End If
			End If
		Else
			If ConditionalMode Then
				PrintMigrateUser = True
				SetLogDir "Conditional", "PrintMig", strUserDept
			End If
			If InteractiveMode Then
				SetLogDir "Interactive", "PrintMig", strUserDept
			End If
		End If
	End If
Else
	Set objShell = nothing
	Set objFSO = nothing
	Set objWMIService = nothing
	Set objNetwork = nothing
	Set objRecordSet = nothing
	Set objCommand = nothing
	objConnection.Close
	Set objConnection = nothing
	WScript.Quit
End If



If SinglePrinterTask or WipeAllPrintersTask Then
	PrintMigrateUser = False
	If SinglePrinterTask Then
		SetLogDir "AutoSelection", "SingleTask", strUserDept
	Else
		SetLogDir "WipeAll", "WipeTask", strUserDept
	End If
	MigrationPreparation()

	If ConditionalMode Then 
		If CustomGroup = vbNullString Then
			objOutput.writeline "Validating user against Department " & strUserDept & " default Printer Migration Group and verifying Mapping file :"
			objOutput.writeline
			If PrintMigGroupError = vbNullString or PrintMigGroupError = " recursively" Then
				ObjOutput.writeline "User " & strUserName & PrintMigGroupError & " validated as a member of default Printer Migration Group " & CHR(34) & strUserDept & "-PrinterMigration" & CHR(34)
			Else
				objOutput.writeline PrintMigGroupValidationLogLineError
			End If
		Else
			objOutput.writeline "Validating user against Department " & strUserDept & " custom Printer Migration Group :"
			objOutput.writeline
			ObjOutput.writeline "User " & strUserName & " validated as a member of custom Printer Migration Group " & CHR(34) & CustomGroup & CHR(34)
		End If
		objOutput.writeline "---------------------------------------------------------------------------------------------------"
		objOutput.writeline
	End If
End If


'Single Task Mode: Carry the Single Task action based on info supplied in the single or two arguments passed on the command line
If SinglePrinterTask Then
	netPrinters = getNetPrinters()
	netPrintersCount = UBound(netPrinters) + 1
	
	If PrinterToAffix <> vbNullString Then
		objOutput.writeline
		If ConditionalMode Then 
			objOutput.writeline "Conditional single printer Affix task status :"
		Else
			objOutput.writeline "Single printer Affix task status :"
		End If
		objOutput.writeline
		AffixPrinterToAffix (PrinterToAffix)
	End If
	
	If PrinterToRemove <> vbNullString Then
		objOutput.writeline
		If ConditionalMode Then 
			objOutput.writeline "Conditional single printer Removal task status :"
		Else
			objOutput.writeline "Single printer Removal task status :"
		End If
		objOutput.writeline
		RemovePrinterToRemove(PrinterToRemove)
	End If
End If


'Wipe/remove all printers Task Mode: used for instance during Department move to new building for instance to prevent users from carrying over printers from previous premisses
If WipeAllPrintersTask Then
	netPrinters = getNetPrinters()
	netPrintersCount = UBound(netPrinters) + 1
	If netPrinters(0) <> "No printers found" and netPrinters(0) <> "Cannot enumerate printers" Then
		objOutput.writeline
		If ConditionalMode Then 
			objOutput.writeline "Conditional All Printers Removal task status :"
		Else
			objOutput.writeline "All Printers Removal task status :"
		End If
		objOutput.writeline
		For j = 0 to UBound(netPrinters)
			strCurrentPrinterShortUNC = Replace(UCase(netPrinters(j)),",DEFAULT",vbNullString)
			PrinterToRemove = Replace(strCurrentPrinterShortUNC,CHR(46) & strDomain,vbNullString)
			objOutput.writeline "Removing installed printer " & PrinterToRemove
			If RemovePrinter(strCurrentPrinterShortUNC) Then
				RemovedPrinters = RemovedPrinters + 1
			Else
				UnremovablePrinters = UnremovablePrinters + 1
			End If
		Next
	End If
End If


If SinglePrinterTask or WipeAllPrintersTask Then
	MigrationSummary()
	MigrationClosure()
End If


'Printer Migration Mode: Carry the Printer Migration based on mapping information located in the Department's old to new print queue Mappings Table file.
If PrintMigrateUser = True and (PrintMigGroupError = vbNullString or PrintMigGroupError = " recursively" or PrintMigGroupError = "NoGroupCheck" or Left(PrintMigGroupValidationLogLineError,8) = "Warning:") Then
	MigrationPreparation()
	
	If ConditionalMode Then
		If CustomGroup = vbNullString Then
			objOutput.writeline "Validating user against Department " & strUserDept & " default Printer Migration Group and verifying Mapping file :"
			objOutput.writeline
			If PrintMigGroupError = vbNullString or PrintMigGroupError = " recursively" Then
				ObjOutput.writeline "User " & strUserName & PrintMigGroupError & " validated as a member of default Printer Migration Group " & CHR(34) & strUserDept & "-PrinterMigration" & CHR(34)
			Else
				objOutput.writeline PrintMigGroupValidationLogLineError
			End If
		Else
			objOutput.writeline "Validating user against Department " & strUserDept & " custom Printer Migration Group and verifying Mapping file :"
			objOutput.writeline
			ObjOutput.writeline "User " & strUserName & PrintMigGroupError & " validated as a member of custom Printer Migration Group " & CHR(34) & CustomGroup & CHR(34)
		End If
		objOutput.writeline
	Else
		If PrintMigGroupError = vbNullString or PrintMigGroupError = " recursively" Then
			objOutput.writeline "Validating Printer Migration Group and Mapping file for Department " & strUserDept & " :"
		Else
			objOutput.writeline "Validating Mapping file for Department " & strUserDept & " :"
		End If
		objOutput.writeline
	End If

	'Get and check the correspondence/mapping list between old and new printers and store it in variable MappingTable. If file is validated/OK then start migrating printers otherwise abort migration and report the mapping file error.
	If objFSO.FileExists(MappingFile) Then
		Set objMappingFile = objFSO.OpenTextFile(MappingFile, ForReading)
		MappingTable = objMappingFile.ReadAll
		AllMappings = Split(MappingTable, vbCRLF)
		objMappingFile.Close
		Set objMappingFile = nothing

		For i = 0 to UBound(AllMappings)
			Select Case Len(AllMappings(i))-Len(Replace(AllMappings(i),CHR(32),vbNullString))
				Case 0
				Case 1
					If Len(Replace(AllMappings(i),CHR(32),vbNullString)) < Len(Trim(AllMappings(i))) Then
						objOutput.writeline "Warning : Mapping file " & MappingFileName & " line " & (i+1) & " contains one space character."
						If MappingFileWarning = vbNullString Then
							MappingFileWarning = "Printer mapping file contains at least one line with space characters! Please trim file!"
						End If
					End If
				Case Else
					If Len(Replace(AllMappings(i),CHR(32),vbNullString)) < Len(Trim(AllMappings(i))) Then
						objOutput.writeline "Warning : Mapping file " & MappingFileName & " line " & (i+1) & " contains " & Len(Trim(AllMappings(i))) - Len(Replace(AllMappings(i),CHR(32),vbNullString)) & " space characters."
						If MappingFileWarning = vbNullString Then
							MappingFileWarning = "Printer mapping file contains at least one line with space characters! Please trim file!"
						End If
					End If
			End Select
			
			'Read from the mapping file's first line if printer migrations are allowed under Terminal Services.
			If i = 0 and Instr(AllMappings(0),CHR(61)) Then
				If UCase(Mid(AllMappings(0),1,Instr(AllMappings(0),CHR(61))-1)) = "ALLOWTSMIGRATIONS" and UCase(Mid(AllMappings(0),Instr(AllMappings(0),CHR(61))+1)) = "FALSE" Then
					PrintMigUnderTS = False
				End If
			Else
				If Instr(AllMappings(i),CHR(61)) Then
					objOutput.writeline "Error : Mapping file " & MappingFileName & " line " & (i+1) & " contains an illegal equal sign."
					MappingFileError = "Printer mapping file contains an equal sign on an other line than the first line. No printer migration performed! Please check file!"
				End If
			End If
			
			If AllMappings(i) <> vbNullString and InStr(AllMappings(i),CHR(44)) and Left(AllMappings(i),2) = "\\" Then
				'This check looks if one or more printers is/are defined to be migrated to itself/themselves.
				arrOldToNewPrinter = split(AllMappings(i),CHR(44))
				strOldPrinterUNC = UCase(arrOldToNewPrinter(0))
				strNewPrinterUNC = UCase(arrOldToNewPrinter(1))
				
				If strOldPrinterUNC = strNewPrinterUNC Then
					If InconsistentLines = vbNullString Then
						InconsistentLines = "Error : Source printer " & strOldPrinterUNC & " is defined to be migrated to itself in " & MappingFileName & " line " & (i+1)
					Else
						InconsistentLines = InconsistentLines  & vbCRLF & "Error : Source printer " & strOldPrinterUNC & " is defined to be migrated to itself in " & MappingFileName & " line " & (i+1)
					End If
					InconsistentLineCount = InconsistentLineCount + 1
					If InconsistentLineCount = 1 Then
						MappingFileError = "Printer mapping file contains one line where a printer is defined to be migrated to itself. No printer migration performed! Please check file!"
					Else
						MappingFileError = "Printer mapping file contains a total of " & InconsistentLineCount & " lines with printers defined to be migrated to themselves. No printer migration performed! Please check file!"
					End If
					If InStr(InconsistencyTracker,"A") = 0 Then InconsistencyTracker = InconsistencyTracker & "A"
				End If
				
				'Detect if the mapping file contains duplicate mappings (source-source or source-destination) for the same source printer
				If i < UBound(AllMappings) Then
					For ii = i + 1 to UBound(AllMappings)
						If AllMappings(ii) <> vbNullString and Instr(AllMappings(ii),CHR(44)) and Left(AllMappings(ii),2) = "\\" Then
							If UCase(Mid(AllMappings(ii),1,Instr(AllMappings(ii),CHR(44))-1)) = strOldPrinterUNC Then
								If InconsistentLines = vbNullString Then
									InconsistentLines = "Error : Source printer " & strOldPrinterUNC & " is referenced twice in mapping file " & MappingFileName & " (lines " & (i+1) & " and " & (ii+1) & ")."
								Else
									InconsistentLines = InconsistentLines  & vbCRLF & "Error : Source printer " & strOldPrinterUNC & " is referenced twice in mapping file " & MappingFileName & " (lines " & (i+1) & " and " & (ii+1) & ")."
								End If
								InconsistentLineCount = InconsistentLineCount + 1
								If InconsistentLineCount = 1 Then
									MappingFileError = "Printer mapping file contains one duplicate entry for the same source printer. No printer migration performed! Please check file!"
								Else
									MappingFileError = "Printer mapping file contains a total of " & InconsistentLineCount & " duplicate entries. No printer migration performed! Please check file!"
								End If
								If InStr(InconsistencyTracker,"B") = 0 Then InconsistencyTracker = InconsistencyTracker & "B"
							Else
								If UCase(Mid(AllMappings(ii),Instr(AllMappings(ii),CHR(44))+1)) <> "INSTALL" Then
									If strOldPrinterUNC = UCase(Mid(AllMappings(ii),Instr(AllMappings(ii),CHR(44))+1)) or strNewPrinterUNC = UCase(Mid(AllMappings(ii),1,Instr(AllMappings(ii),CHR(44))-1)) Then
										If strNewPrinterUNC = UCase(Mid(AllMappings(ii),1,Instr(AllMappings(ii),CHR(44))-1)) Then
												If InconsistentLines = vbNullString Then
													InconsistentLines = "Error : Destination printer " & strNewPrinterUNC & " is also referenced in mapping file " & MappingFileName & " as source printer (lines " & (i+1) & " as destination and " & (ii+1) & " as source)."
												Else
													InconsistentLines = InconsistentLines  & vbCRLF & "Error : Destination printer " & strNewPrinterUNC & " is also referenced in mapping file " & MappingFileName & " as source printer (lines " & (i+1) & " as destination and " & (ii+1) & " as source)."
												End If
										Else
											If InconsistentLines = vbNullString Then
												InconsistentLines = "Error : Source printer " & strOldPrinterUNC & " is also referenced in mapping file " & MappingFileName & " as destination printer (lines " & (i+1) & " as source and " & (ii+1) & " as destination)."
											Else
												InconsistentLines = InconsistentLines  & vbCRLF & "Error : Source printer " & strOldPrinterUNC & " is also referenced in mapping file " & MappingFileName & " as a destination printer (lines " & (i+1) & " as source and " & (ii+1) & " as destination)."
											End If
										End If
										InconsistentLineCount = InconsistentLineCount + 1
										If InconsistentLineCount = 1 Then
											MappingFileError = "Printer mapping file contains one printer which is defined both as source and as destination. No printer migration performed! Please check file!"
										Else
											MappingFileError = "Printer mapping file contains a total of " & InconsistentLineCount & " printers which are defined both as source and as destination. No printer migration performed! Please check file!"
										End If
										If InStr(InconsistencyTracker,"C") = 0 Then InconsistencyTracker = InconsistencyTracker & "C"
									Else
										If Len(Replace(strOldPrinterUNC,CHR(92),vbNullString)) = Len(strOldPrinterUNC) - 2 and strNewPrinterUNC = "DELETE" and InStr(UCase(Mid(AllMappings(ii),1,Instr(AllMappings(ii),CHR(44))-1)),strOldPrinterUNC) and Len(Replace(Mid(AllMappings(ii),Instr(AllMappings(ii),CHR(44))+1),CHR(92),vbNullString)) = Len(Mid(AllMappings(ii),Instr(AllMappings(ii),CHR(44))+1)) - 3 Then
											If InconsistentLines = vbNullString Then
												InconsistentLines = "Error : Entry for obsolete Server " & strOldPrinterUNC & " in mapping file " & MappingFileName & " line " & (i+1) & " cannot be before migration mapping line " & (ii+1) & " for Printer " & UCase(Mid(AllMappings(ii),1,Instr(AllMappings(ii),CHR(44))-1)) & " on this same server."
											Else
												InconsistentLines = InconsistentLines  & vbCRLF & "Error : Entry for obsolete Server " & strOldPrinterUNC & " in mapping file " & MappingFileName & " line " & (i+1) & " cannot be before migration mapping line " & (ii+1) & " for Printer " & UCase(Mid(AllMappings(ii),1,Instr(AllMappings(ii),CHR(44))-1)) & " on this same server."
											End If
											InconsistentLineCount = InconsistentLineCount + 1
											MappingFileError = "Printer mapping file contains a server delete line before a migration line for a printer on this same server. No printer migration performed! Please check file!"
											If InStr(InconsistencyTracker,"C") = 0 Then InconsistencyTracker = InconsistencyTracker & "C"
										End If
									End If
								End If
							End If
						End If
					Next
				End If
				
				'Mapping file check for number of commas added
				If Len(AllMappings(i)) - Len(Replace(AllMappings(i),CHR(44),vbNullString)) > 1 Then
					InconsistentLineCount = InconsistentLineCount + 1
					If InconsistentLines = vbNullString Then
						InconsistentLines = "Error : Printer mapping file " & MappingFileName & " line " & i+1 & " contains " & Len(AllMappings(i)) - Len(Replace(AllMappings(i),CHR(44),vbNullString)) & " commas instead of only one."
					Else
						InconsistentLines = InconsistentLines  & vbCRLF & "Error : Printer mapping file " & MappingFileName & " line " & i+1 & " contains " & Len(AllMappings(i)) - Len(Replace(AllMappings(i),CHR(44),vbNullString)) & " commas instead of only one."
					End If
					If InconsistentLineCount = 1 Then
						MappingFileError = "Printer mapping file contains one line with more than one comma. No printer migration performed!" & vbCRLF & "Please check and correct mapping file " & MappingFile & "!" 
					Else
						MappingFileError = "Printer mapping file contains " & InconsistentLineCount & " lines with more than one comma. No printer migration performed!" & vbCRLF & "Please check and correct mapping file " & MappingFile & "!" 
					End If
					If InStr(InconsistencyTracker,"D") = 0 Then InconsistencyTracker = InconsistencyTracker & "D"
				End If
				If InStr(UCase(AllMappings(i)),",INSTALL") Then
					If PrintersToAffix = vbNullString Then
						PrintersToAffix = strOldPrinterUNC
					Else
						PrintersToAffix = PrintersToAffix & CHR(124) & strOldPrinterUNC
					End If
				End If
			Else
				If AllMappings(i) <> vbNullString and Len(AllMappings(i)) > 15 and Left(AllMappings(i),2) = "\\" and InStr(AllMappings(i),CHR(44)) = 0 Then
					If InStr((Mid(AllMappings(i),3)),"\\") Then
						CurrentFieldSeparator = Mid(Mid(AllMappings(i),3),InStr((Mid(AllMappings(i),3)),"\\")-1,1)
						If AllFieldSeparators = vbNullString Then
							AllFieldSeparators = CurrentFieldSeparator
						Else
							If InStr(AllFieldSeparators,CurrentFieldSeparator) = 0 Then 
								AllFieldSeparators = AllFieldSeparators & CHR(32) & CurrentFieldSeparator
							End If
						End If
						InvalidSeparatorLineCount = InvalidSeparatorLineCount + 1
						If InvalidSeparatorLineCount < Int(UBound(AllMappings)/2) Then
							If InconsistentLines = vbNullString Then
								InconsistentLines = "Error : Printer mapping file " & MappingFileName & " line " & i+1 & " is using '" & CurrentFieldSeparator & "' as a field separator instead of a comma."
							Else
								InconsistentLines = InconsistentLines  & vbCRLF &  "Error : Printer mapping file " & MappingFileName & " line " & i+1 & " is using '" & CurrentFieldSeparator & "' as a field separator instead of a comma."
							End If
							If InvalidSeparatorLineCount = 1 Then
								MappingFileError = "Printer mapping file contains one line with an invalid field separator ('" & CurrentFieldSeparator& "' instead of ',') . No printer migration performed!" & vbCRLF & "Please check and correct mapping file " & MappingFile & "!"
							Else
								MappingFileError = "Printer mapping file contains " & InvalidSeparatorLineCount & " lines with invalid field separators ('" & AllFieldSeparators& "' instead of ',') . No printer migration performed!" & vbCRLF & "Please check and correct mapping file " & MappingFile & "!"
							End If
						Else
							If InconsistentLines = vbNullString Then
								InconsistentLines = "Error : Printer mapping file " & MappingFileName & " is using '" & AllFieldSeparators & "' as seperator instead of commas."
							Else
								InconsistentLines = InconsistentLines  & vbCRLF &  "Error : Printer mapping file " & MappingFileName & " is using '" & AllFieldSeparators & "' as seperator instead of commas."
							End If
							MappingFileError = "Printer mapping file is using '" & AllFieldSeparators & "' as seperator instead of commas. No printer migration performed!" & vbCRLF & "Please check and correct mapping file " & MappingFile & "!" 
							Exit For
						End If
						If InStr(InconsistencyTracker,"E") = 0 Then InconsistencyTracker = InconsistencyTracker & "E"
					Else
						If InconsistentLines = vbNullString Then
							InconsistentLines = "Error : Printer mapping file " & MappingFileName & " line " & i+1 & " is invalid."
						Else
							InconsistentLines = InconsistentLines  & vbCRLF &  "Error : Printer mapping file " & MappingFileName & " line " & i+1 & " is invalid."
						End If
						MappingFileError = "Printer mapping file contains at least one invalid line. No printer migration performed!" & vbCRLF & "Please check and correct mapping file " & MappingFile & "!" 
					End If
				Else
					'Inconsistent line check
					If AllMappings(i) <> CHR(13) and Len(AllMappings(i)) > 0 and Left(AllMappings(i),1) <> CHR(32) Then
						If MappingFileError = vbNullString and InStr(AllMappings(0),CHR(61)) = 0 Then
							objOutput.writeline "Error : Printer mapping file " & MappingFileName & " line " & i+1 & " is inconsistent."
							MappingFileError = "Printer mapping file has at least one inconsistent line. No printer migration performed!" & vbCRLF & "Please check and correct mapping file " & MappingFile & "!" 
							Exit For
						End If
					End If
				End If
				'Mapping file field separator check added
				If Left(AllMappings(i),1) = CHR(34) and Right(AllMappings(i),1) = CHR(34) Then
					objOutput.writeline "Error : Printer mapping file " & MappingFileName & " has its mappings enclosed between double quotes."
					MappingFileError = "Printer mapping file has its mappings enclosed between double quotes. No printer migration performed!" & vbCRLF & "Please check and correct mapping file " & MappingFile & "!" 
					Exit For
				End If
			End If
		Next
		
		If PrintMigUnderTS = False and TSSession = True Then
			Set objShell = nothing
			objOutput.Close
			Set objOutput = nothing
			objFSO.DeleteFile(UserLogFile)
			Set objFSO = nothing
			Set objWMIService = nothing
			Set objNetwork = nothing
			Set objCommand = nothing
			objConnection.Close
			Set objConnection = nothing
			WScript.Quit
		Else
			If MappingFileError = vbNullString Then
				If i = 0 Then
					objOutput.writeline "Error : Printer mapping file " & MappingFileName & " appears to be corrupted."
					MappingFileError = "Printer mapping file corrupted. No printer migration performed!" & vbCRLF & "Please check and correct mapping file " & MappingFile & "!"
				Else
					If PrintMigGroupError = vbNullString or PrintMigGroupError = " recursively" Then
						If ConditionalMode Then
							If MappingFileWarning = vbNullString Then
								objOutput.writeline "Printer Mapping file " & MappingFileName & " passed all checks!"
							Else
								objOutput.writeline "Printer Mapping file " & MappingFileName & " passed all checks but with warnings!"
							End If
						Else
							If MappingFileWarning = vbNullString Then
								objOutput.writeline "Group " & strUserDept & "-PrinterMigration exists and Printer Mapping file " & MappingFileName & " passed all checks!"
							Else
								objOutput.writeline "Group " & strUserDept & "-PrinterMigration exists and Printer Mapping file " & MappingFileName & " passed all checks but with warnings!"
							End If
						End If
					Else
						If MappingFileWarning = vbNullString Then
							objOutput.writeline "Printer Mapping file " & MappingFileName & " passed all checks!"
						Else
							objOutput.writeline "Printer Mapping file " & MappingFileName & " passed all checks but with warnings!"
						End If
						objOutput.writeline PrintMigGroupValidationLogLineError
					End If
				End If
			Else
				If InconsistentLines <> vbNullString Then
					objOutput.writeline InconsistentLines
					If Len(InconsistencyTracker) > 1 Then
						MappingFileError = "Printer mapping file contains several inconsistencies. No printer migration performed!" & vbCRLF & "Please check and correct mapping file " & MappingFile & "!"
					End If
				End If
			End If
		End If
	Else
		If MappingFileError = vbNullString Then
			objOutput.writeline "Error : Printer mapping file " & MappingFileName & " doesn't exist or cannot be accessed!"
			MappingFileError = "Printer mapping file " & MappingFileName & " missing. No printer migration performed! " & vbCRLF & "Please create and populate mapping file " & MappingFile & "!"
		Else
			objOutput.writeline "Error : Printer mapping file " & MappingFileName & " specified interactively is not a valid mapping file name!"
		End If
	End If
	InconsistentLines = Null

	If MappingFileError = vbNullString Then
		objOutput.writeline "---------------------------------------------------------------------------------------------------"
		objOutput.writeline
		netPrinters = getNetPrinters()
		
		netPrintersCount = UBound(netPrinters) + 1
		If netPrinters(0) <> "No printers found" and netPrinters(0) <> "Cannot enumerate printers" Then
			objOutput.writeline
			objOutput.writeline "New network printers and migration status :"
			objOutput.writeline
			MigratePrinters()
		End If

		If PrintersToAffix <> vbNullString Then
			If netPrinters(0) = "No printers found" Then
				objOutput.writeline "---------------------------------------------------------------------------------------------------"
				objOutput.writeline
				objOutput.writeline "New network printers status :"
			End If
			objOutput.writeline
			AffixedPrinters = AffixCompulsoryPrinters(PrintersToAffix)
		End If
		Erase netPrinters
	End If

	MigrationSummary()
	MigrationClosure()
End If


'End cleanup and object destructions
Erase TimeTrack
Set objShell = nothing
Set objFSO = nothing
Set objWMIService = nothing
Set objNetwork = nothing
Set objRecordSet = nothing
Set objCommand = nothing
objConnection.Close
Set objConnection = nothing
WScript.Quit



'End of script Main body. Functions and SubRoutines declared below
'####################################################################################################



Function DeptPrintMigGroupExists (Dept)
	Dim strBase
	Dim strLDAPQuery
	Dim strGroupDN, arrGroupDN, strGroupADPath
	Dim k

	If Dept <> "OTHER" Then
		'Search for presence of Department specific Printer Migration group in the IT OU'
		strBase = PrintMigGroupsOU & strDomainDN & CHR(62)  
		strLDAPQuery = strBase & CHR(59) & "(&(objectCategory=group)(Name=" & Dept & "-PrinterMigration" & "));distinguishedName;subtree"
		objCommand.CommandText = strLDAPQuery
		Set objRecordSet = objCommand.Execute 
		If Not objRecordset.EOF Then 
			DeptPrintMigGroupExists = True
		Else
			'Could not locate Group in expected IT OU (as specified in 'PrintMigGroupsOU' configurable variable), look for this group in entire AD scope
			strBase = "<LDAP://" & strDomainDN & CHR(62)
			strLDAPQuery = strBase & CHR(59) & "(&(objectCategory=group)(Name=" & Dept & "-PrinterMigration" & "));distinguishedName;subtree"
			objCommand.CommandText = strLDAPQuery
			Set objRecordSet = objCommand.Execute
			If Not objRecordset.EOF Then
				DeptPrintMigGroupExists = True
				strGroupDN = objRecordSet.Fields("distinguishedName").Value
				arrGroupDN = split(strGroupDN,CHR(44))
				strGroupADPath = vbNullString
				For k = UBound(arrGroupDN) to 0 Step -1
					If InStr(arrGroupDN(k),"OU=") Then
						strGroupADPath = strGroupADPath & CHR(47) & arrGroupDN(k)
					End If
				Next
				PrintMigGroupValidationLogLineError = "Warning: Printer Migration group " & strUserDept & "-PrinterMigration for Department " & strUserDept & " resides in the wrong AD OU container!" & vbCRLF & "Warning: Printer Migration group " & strUserDept & "-PrinterMigration for Department " & strUserDept & " is wrongfully published in " & strGroupADPath
				PrintMigGroupError = "The Printer Migration group " & strUserDept & "-PrinterMigration for Department " & strUserDept & " resides in the wrong AD OU container! Please move this group to the configured '" & PrintMigGroupsOU & "' OU or specify it as a Custom group on the script's command line!"
			Else
				DeptPrintMigGroupExists = False
				PrintMigGroupValidationLogLineError = "Error: Printer Migration group " & strUserDept & "-PrinterMigration for Department " & strUserDept & " does not exist!"
				PrintMigGroupError = "The Printer Migration group " & strUserDept & "-PrinterMigration for Department " & strUserDept & " does not exist! Please create the group and populate it with users to migrate!"
			End If
		End If
	Else
		DeptPrintMigGroupExists = True
	End If
End Function
'----------------------------------------------------------------------------------------------------



Function GroupMember (GroupName)
	Dim objUser, objGroup
	Dim colUserMemberOf, colUserGroups
	
	GroupMember = False
	Set objUser = GetObject("LDAP://" & strUserDN)
	
	On Error Resume Next
	colUserMemberOf = objUser.GetEx("memberOf")
	If Err.Number = 0 Then
		For Each objGroup in colUserMemberOf
			If UCase(Mid(objGroup,4,InStr(objGroup,CHR(44))-4)) = UCase(GroupName) Then
				GroupMember = True
				On Error Goto 0
				Exit Function
			End If
		Next
	Else
		Err.Clear
		On Error Goto 0
		Exit Function
	End If

	Set colUserGroups = objUser.Groups
	For Each objGroup in colUserGroups
		If UCase(objGroup.CN) <> UCase(GroupName) Then
			If UserGroupList = vbNullString Then
				UserGroupList = objGroup.CN
			Else
				arrUserGroupList = split(UserGroupList,CHR(124))
				If UCase(arrUserGroupList(UBound(arrUserGroupList))) <> UCase(GroupName) Then
					UserGroupList = UserGroupList & CHR(124) & objGroup.CN
				Else
					Exit For
				End If
			End If
			SearchNestedGroup GroupName,objGroup
		End If
	Next
	If UCase(arrUserGroupList(UBound(arrUserGroupList))) = UCase(GroupName) Then
		GroupMember = True
		PrintMigGroupError = " recursively"
	End If

	On Error Goto 0
	Set objUser = nothing
End Function
'----------------------------------------------------------------------------------------------------



Function SearchNestedGroup(GroupName,objParentGroup)
	Dim GroupMemberOf, colGroupMemberOf
	Dim GroupAlreadyListed
	Dim objNestedGroup
	Dim UserGroup
	
	On Error Resume Next
	colGroupMemberOf = objParentGroup.GetEx("memberOf")
	If Err.Number = 0 Then
		For Each GroupMemberOf in colGroupMemberOf
			GroupAlreadyListed = False
			Set objNestedGroup = GetObject("LDAP://" & GroupMemberOf)
			If Err.Number = 0 Then
				If UCase(objNestedGroup.CN) = UCase(GroupName) Then
					UserGroupList = UserGroupList & CHR(124) & objNestedGroup.CN
					On Error Goto 0
					Exit Function
				Else
					arrUserGroupList = split(UserGroupList,CHR(124))
					For Each UserGroup in arrUserGroupList
						If objNestedGroup.CN = UserGroup Then
							GroupAlreadyListed = True
							Exit For
						End If
					Next
					If GroupAlreadyListed = False Then
						UserGroupList = UserGroupList & CHR(124) & objNestedGroup.CN
						SearchNestedGroup GroupName,objNestedGroup
					End If
				End If
			Else
				Err.Clear
			End If
			arrUserGroupList = split(UserGroupList,CHR(124))
			If UCase(arrUserGroupList(UBound(arrUserGroupList))) = UCase(GroupName) Then
				On Error Goto 0
				Exit Function
			End If
		Next
	Else
		Err.Clear
	End If
	
	On Error Goto 0
End Function
'----------------------------------------------------------------------------------------------------



Sub	SetLogDir (SelectionMode,OperationType,Dept)
	Dim ModeDir
	Dim NewLogModeDir
	
	Select Case SelectionMode
		Case "Interactive"
			ModeDir = "Interactive\"
		Case "Conditional"
			ModeDir = "Conditional\"
		Case "WipeAll"
			ModeDir = "WipeAll\"
		Case Else
			ModeDir = vbNullString
	End Select
	
	OperationType = OperationType & CHR(45)
	
	If objFSO.FolderExists(PrintQMigrationLogs & ModeDir & OperationType & Dept) Then
		DestLogDir = PrintQMigrationLogs & ModeDir & OperationType & Dept
	Else
		On Error Resume Next
		Set NewLogModeDir = objFSO.CreateFolder(PrintQMigrationLogs & ModeDir & OperationType & Dept)
		If Err.Number Then
			Err.Clear
			If objFSO.FolderExists(PrintQMigrationLogs & ModeDir & OperationType & "OTHER") Then
				DestLogDir = PrintQMigrationLogs & ModeDir & OperationType & "OTHER"
			Else
				Set NewLogModeDir = objFSO.CreateFolder(PrintQMigrationLogs & ModeDir & OperationType & "OTHER")
				If Not Err.Number Then
					DestLogDir = PrintQMigrationLogs & ModeDir & OperationType & "OTHER"
				Else
					Err.Clear
					Set NewLogModeDir = nothing
					Set objOutput = nothing
					Set objShell = nothing
					Set objFSO = nothing
					Set objWMIService = nothing
					Set objNetwork = nothing
					Set objRecordSet = nothing
					Set objCommand = nothing
					objConnection.Close
					Set objConnection = nothing
					WScript.Quit
				End If
			End If
		Else
			DestLogDir = PrintQMigrationLogs & ModeDir & OperationType & Dept
		End If
		On Error Goto 0
		Set NewLogModeDir = nothing
	End If
End Sub
'----------------------------------------------------------------------------------------------------



Sub MigrationPreparation
	Dim objReg
	Dim objExec, RDPConnection

	'User Printer migration logs
	UserLogTimeStamp= Year(Now) & CHR(45) & Right("0" & Month(Now),2) & CHR(45) & Right("0" & Day(Now),2) & CHR(95) & Right("0" & Hour(Now),2) & CHR(45) & Right("0" & Minute(Now),2) & CHR(45) & Right("0" & Second(Now),2)
	If SinglePrinterTask = False and WipeAllPrintersTask = False Then
		UserLogFile = DestLogDir & CHR(92) & "PrintMig_" & strUserDept & CHR(95) & lCase(strUserName) & CHR(95) & strComputerName & CHR(95) & UserLogTimeStamp & ".log"
	Else
		If WipeAllPrintersTask Then
			UserLogFile = DestLogDir & CHR(92) & "WipeAll_" & strUserDept & CHR(95) & lCase(strUserName) & CHR(95) & strComputerName & CHR(95) & UserLogTimeStamp & ".log"
		Else
			UserLogFile = DestLogDir & CHR(92) & "SingleTask_" & strUserDept & CHR(95) & lCase(strUserName) & CHR(95) & strComputerName & CHR(95) & UserLogTimeStamp & ".log"
		End If
	End If
		
	'Start Printer Migration for users (group membership validated in parent/caller Logon Script) and group membership validated in present script in case of non Logon Script invocation users
	'Create the user's log file. As this can sometimes fail on DFS shares, abort script operation
	On Error Resume Next
	Set objOutput = objFSO.OpenTextFile(UserLogFile, ForWriting, True)
	If Err.Number Then
		Err.Clear
		Set objOutput = nothing
		Set objShell = nothing
		Set objFSO = nothing
		Set objWMIService = nothing
		Set objNetwork = nothing
		Set objRecordSet = nothing
		Set objCommand = nothing
		objConnection.Close
		Set objConnection = nothing
		WScript.Quit
	End If
	On Error Goto 0

	Set objReg = GetObject("winmgmts:!root/default:StdRegProv")
	objReg.GetStringValue HKEY_LOCAL_MACHINE,"SOFTWARE\Microsoft\Windows NT\CurrentVersion","ProductName",OSVersion
	
	'This next block of code (next 12 lines) detects if the current session is a Terminal Server/Remote Desktop Service session. This method is preferred over grabbing the SESSIONNAME volatile environment variable since it is not yet populated when logging on under Wintows 7 and TS 2008.
	Set objExec = objShell.Exec("cmd /c QWInsta |find """ & strUserName & """ /i |find ""Active""") 
	RDPConnection = objExec.StdOut.ReadLine
	If Left(RDPConnection,8) = ">console" Then
		TSSession = False
	End If
	If Left(RDPConnection,9) = ">rdp-tcp#" Then
		TSSession = True
	End If
	Set objExec = nothing
	If TSSession = vbNullString Then
		objOutput.writeline "Could not determine Terminal Session mode with Printer Management script executed on " & strStationType & CHR(32) & strComputerName & " running " & OSVersion & " for user " & strUserName
	End If

	If UCase(Left(strComputerName,1)) = "L" Then
		strStationType = "20" & Mid(strComputerName,6,2) & CHR(32) & ComputerBrand & "laptop"
	End If
	If UCase(Left(strComputerName,1)) = "D" Then
		strStationType = "20" & Mid(strComputerName,6,2) & CHR(32) & ComputerBrand & "desktop"
	End If
	If InStr(lCase(OSVersion),"server") Then
		strStationType = "terminal server"
	End If
	If InteractiveMode = True Then
		objOutput.writeline "Client Printer Management script executed interactively on " & strStationType & CHR(32) & strComputerName & " running " & OSVersion & " for user " & strUserName
	Else
		objOutput.writeline "Client Printer Management script executed on " & strStationType & CHR(32) & strComputerName & " running " & OSVersion & " for user " & strUserName
	End If
	objOutput.writeline
	objOutput.writeline
	Set objReg = nothing

	If strUserDept = "OTHER" Then
		objOutput.writeline "User " & strUserName & " was not found in a valid OU Container (i.e. no valid Department)."
		objOutput.writeline "The User's AD Distinguished Name (AD path) is :"
		objOutput.writeline "	" & strUserDN
		objOutput.writeline "---------------------------------------------------------------------------------------------------"
		objOutput.writeline
	End If
End Sub
'----------------------------------------------------------------------------------------------------



Function ComputerBrand ()
	Dim colAdapters, objAdapter
	
	Set colAdapters = objWMIService.ExecQuery ("SELECT * FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled = True")
	For Each objAdapter in colAdapters
		Select Case Left(objAdapter.MACAddress,8)
			Case "00:19:99"
				ComputerBrand = "Fujitsu Siemens "
			Case "B8:AC:6F"
				ComputerBrand = "Dell "
			Case "00:22:19"
				ComputerBrand = "Dell "
			Case "00:25:64"
				ComputerBrand = "Dell "
			Case "00:1D:09"
				ComputerBrand = "Dell "
			Case "00:1E:C9"
				ComputerBrand = "Dell "
			Case "18:A9:05"
				ComputerBrand = "Hewlett Packard "
			Case "D8:D3:85"
				ComputerBrand = "Hewlett Packard "
			Case Else
				ComputerBrand = vbNullString
		End Select
	Next
	Set colAdapters = nothing
End Function
'----------------------------------------------------------------------------------------------------



Function getNetPrinters ()
	Dim iPrinter
	Dim arrWMIPrinters
	Dim arrWSHPrinters
	Dim arrPrinters()
	Dim CheckEmpty
	Dim InstalledPrinterLine
	Dim BlankStream
	Dim iSpaceChar
	Dim IdentifiedPrinterLength
	Dim PrinterEnumerationError : PrinterEnumerationError = vbNullString
	Dim WMIPrintersList, WMIWSHPrintersList
	Dim objFolder, objFolderItem
	Dim colPrinterItems, objPrinterItem, InstalledWinPrinter
	Dim DefaultIdentificationMethod
	Dim IsDefaultPrinter
	Dim objPrinter
	Dim arrSize : arrSize = 0
	Dim InstalledPrintersListShortUNC
	Dim l, m, n
	
	'Printer enumeration method #1: Enumerate installed network printers with WMI and record information in arrPrinters array
	Set arrWMIPrinters = objWMIService.ExecQuery ("SELECT * FROM Win32_Printer WHERE Local = False")
	On Error Resume Next
	If arrWMIPrinters.Count Then
		If Err.Number = 0 Then
			On Error Goto 0
			For Each objPrinter in arrWMIPrinters
				ReDim Preserve arrPrinters(arrSize)
				If objPrinter.Attributes And DEFAULT_PRINTER Then
					DefaultPrinterName = objPrinter.Name
					arrPrinters(arrSize) = "WMI:" & DefaultPrinterName & ",DEFAULT"
				Else
					arrPrinters(arrSize) = "WMI:" & objPrinter.Name
				End If
				arrSize = arrSize + 1
			Next
			Set arrWMIPrinters = nothing
		Else
			PrinterEnumerationError = "Error: Cannot enumerate network printers using WMI!" & vbCRLF
		End If
	End If
	Set arrWMIPrinters = nothing

	'Printer enumeration method #2: Enumerate installed network printers using the WSH Network object and complete the list of printers previously identified by WMI with those it missed
	On Error Resume Next
	WMIPrintersList = Join(arrPrinters,CHR(124))
	Set arrWSHPrinters = objNetwork.EnumPrinterConnections
	If Err.Number = 0 Then
		On Error Goto 0
		For m = 0 to arrWSHPrinters.Count - 1 Step 2
			If Len(Replace(arrWSHPrinters.Item(m+1),CHR(92),vbNullString)) = Len(arrWSHPrinters.Item(m+1)) - 3 Then
				On Error Resume Next
				If Len(WMIPrintersList) > 0 Then
					On Error GoTo 0
					If InStr(UCase(WMIPrintersList),UCase(arrWSHPrinters.Item(m+1))) = 0 Then
						arrSize = UBound(arrPrinters) + 1
						ReDim Preserve arrPrinters(arrSize)
						arrPrinters(arrSize) = "WSH:" & arrWSHPrinters.Item(m+1)
					End If
				Else
					Err.Clear
					On Error GoTo 0
					ReDim Preserve arrPrinters(arrSize)
					arrPrinters(arrSize) = "WSH:" & arrWSHPrinters.Item(m+1)
					arrSize = arrSize + 1
				End If
			End If
		Next
		Set arrWSHPrinters = nothing
	Else
		If Err.Number = 462 Then
			Err.Clear
			On Error Goto 0
			PrinterEnumerationError = PrinterEnumerationError & "Error: Cannot enumerate network printers using the WSH Network object because the Print Spooler service isn't running!" & vbCRLF '& vbCRLF & "Aborting script operation for this logon"
		Else
			On Error Goto 0
			If Err.Description <> vbNullString Then
				PrinterEnumerationError = PrinterEnumerationError & "Error: Enumerating network printers using the WSH Network object returned error " & Err.Number & " (" & Replace(Err.Description,vbCRLF,vbNullString) & " !" & vbCRLF
			Else
				PrinterEnumerationError = PrinterEnumerationError & "Error: Enumerating network printers using the WSH Network object returned error " & Err.Number & " !" & vbCRLF
			End If
			Err.Clear
		End If
		ReDim Preserve arrPrinters(arrSize)
		arrPrinters(arrSize) = "Cannot enumerate printers using WSH"
	End If
	
	'Printer enumeration method #3: Enumerate installed printers from the Printers and Faxes folder using the Shell Application object and complete the list of printers previously identified by WMI and the WSH Network object with those they both missed
	WMI-WSHPrintersList = Join(arrPrinters,CHR(124))
	Set objFolder = objShellApp.Namespace(PRINTERS_AND_FAXES)
	Set objFolderItem = objFolder.Self
	Set colPrinterItems = objFolder.Items
	For Each objPrinterItem in colPrinterItems
		If InStr(objPrinterItem.Name," on ") Then
			InstalledWinPrinter = CHR(92) & CHR(92) & Mid(objPrinterItem.Name,InStrRev(objPrinterItem.Name,CHR(32))+1) & CHR(92) & Mid(objPrinterItem.Name,1,InStr(objPrinterItem.Name,CHR(32)))
			Wscript.Echo objPrinterItem.Name & "<->" & InstalledWinPrinter
			On Error Resume Next
			If Len(WMI-WSHPrintersList) > 0 Then
				On Error GoTo 0
				If InStr(UCase(WMI-WSHPrintersList),UCase(InstalledWinPrinter)) = 0 Then
					arrSize = UBound(arrPrinters) + 1
					ReDim Preserve arrPrinters(arrSize)
					arrPrinters(arrSize) = "WIN:" & InstalledWinPrinter
				End If
			Else
				Err.Clear
				On Error GoTo 0
				ReDim Preserve arrPrinters(arrSize)
				arrPrinters(arrSize) = "WIN:" & InstalledWinPrinter
				arrSize = arrSize + 1
			End If
		End If
	Next
	Set colPrinterItems = nothing
	Set objFolderItem = nothing
	Set objFolder = nothing
	Set objShellApp = nothing
	
	On Error Resume Next
	If UBound(arrPrinters) >= 0 Then
		If Err.Number Then
			Err.Clear
			On Error GoTo 0
			ReDim Preserve arrPrinters(arrSize)
			arrPrinters(arrSize) = "No printers found"
		Else
			On Error GoTo 0
			If DefaultPrinterName = vbNullString Then
				DefaultPrinterName = GetDefaultPrinterFromRegistry
				For iPrinter = 0 to Ubound(arrPrinters)
					If InStr(UCase(arrPrinters(iPrinter)),UCase(DefaultPrinterName)) Then
						arrPrinters(iPrinter) = arrPrinters(iPrinter) & ",DEFAULT-REG"
					End If
				Next
			End If
		End If
	End If

	For iPrinter = 0 to UBound(arrPrinters)
		InstalledPrintersListShortUNC = Replace(UCase(Mid(arrPrinters(iPrinter),5)),CHR(46) & strDomain,vbNullString)
		If InstalledPrintersList = vbNullString Then
			InstalledPrintersList = InstalledPrintersListShortUNC
		Else
			InstalledPrintersList = InstalledPrintersList & CHR(124) & InstalledPrintersListShortUNC
		End If
	Next
	
	If PrinterEnumerationError <> vbNullString  and IsArray(arrPrinters) = False Then
		objOutput.writeline PrinterEnumerationError
		objOutput.writeline "Aborting script operation for this logon"
	Else
		If arrPrinters(0) <> "No printers found" Then
			If UBound(arrPrinters) = 0 Then
				objOutput.writeline "There is 1 network printer currently installed :"
			Else
				objOutput.writeline "There are " & UBound(arrPrinters) + 1 & " network printers currently installed :"
			End If
			objOutput.writeline
			For iPrinter = 0 to UBound(arrPrinters)
				If InStr(arrPrinters(iPrinter),",DEFAULT") Then
					IsDefaultPrinter = True
					If Mid(arrPrinters(iPrinter),len(arrPrinters(iPrinter))-11) = ",DEFAULT-REG" Then
						DefaultIdentificationMethod = "Registry"
						arrPrinters(iPrinter) = Replace(arrPrinters(iPrinter),"-REG",vbNullString)
					Else
						DefaultIdentificationMethod = "WMI"
					End If
				Else
					IsDefaultPrinter = False
				End If

				Select Case Mid(arrPrinters(iPrinter),1,3)
					Case "WMI"
						InstalledPrinterLine = "Printer: " & Mid(Replace(arrPrinters(iPrinter),",DEFAULT"," - DEFAULT"),5)
						BlankStream = vbNullString
						IdentifiedPrinterLength = Len(InstalledPrinterLine)
						For iSpaceChar = IdentifiedPrinterLength to 60
							BlankStream = BlankStream & CHR(32)
						Next
						If IsDefaultPrinter Then
							If DefaultIdentificationMethod = "WMI" Then
								objOutput.writeline InstalledPrinterLine & BlankStream & "[Identified with WMI and Default status also identified with WMI]"
							Else
								objOutput.writeline InstalledPrinterLine & BlankStream & "[Identified with WMI and Default status identified in the Registry]"
							End If
						Else
							objOutput.writeline InstalledPrinterLine & BlankStream & "[Identified with WMI]"
						End If
					Case "WSH"
						InstalledPrinterLine = "Printer: " & Mid(Replace(arrPrinters(iPrinter),",DEFAULT"," - DEFAULT"),5)
						BlankStream = vbNullString
						IdentifiedPrinterLength = Len(InstalledPrinterLine)
						For iSpaceChar = IdentifiedPrinterLength to 60
							BlankStream = BlankStream & CHR(32)
						Next
						If IsDefaultPrinter Then
							objOutput.writeline InstalledPrinterLine & BlankStream & "[Identified with WSH Network object and Default status identified in the Registry]"
						Else
							objOutput.writeline InstalledPrinterLine & BlankStream & "[Identified with WSH Network object]"
						End If
					Case "WIN"
						InstalledPrinterLine = "Printer: " & Mid(Replace(arrPrinters(iPrinter),",DEFAULT"," - DEFAULT"),5)
						BlankStream = vbNullString
						IdentifiedPrinterLength = Len(InstalledPrinterLine)
						For iSpaceChar = IdentifiedPrinterLength to 60
							BlankStream = BlankStream & CHR(32)
						Next
						If IsDefaultPrinter Then
							objOutput.writeline InstalledPrinterLine & BlankStream & "[Identified with the Shell Application object and Default status identified in the Registry]"
						Else
							objOutput.writeline InstalledPrinterLine & BlankStream & "[Identified with the Shell Application object]"
						End If
				End Select
			Next
			objOutput.writeline "---------------------------------------------------------------------------------------------------"
		Else
			objOutput.writeline "No network printers found!"
		End If
	End If
	
	getNetPrinters = arrPrinters
End Function
'----------------------------------------------------------------------------------------------------



Function GetDefaultPrinterFromRegistry
	Dim strDefaultPrinterValue : strDefaultPrinterValue = vbNullString
	Dim objRegistry
	Dim strKeyPath, strValueName
	
	On Error Resume Next
	strDefaultPrinterValue = objShell.RegRead("HKCU\Software\Microsoft\Windows NT\CurrentVersion\Windows\Device")

	If Err.Number Then
		Err.Clear
		On Error GoTo 0
		DumpWindowsProfileInfo "Shell_RegRead","NOK"
	Else
		On Error GoTo 0
		DumpWindowsProfileInfo "Shell_RegRead","OK"
	End If

	strDefaultPrinterValue = UCase(Left(strDefaultPrinterValue,InStr(strDefaultPrinterValue,CHR(44))-1))

	If strDefaultPrinterValue = vbNullString Then
		Set objRegistry = GetObject("winmgmts:!root/default:StdRegProv")
		strKeyPath = "Software\Microsoft\Windows NT\CurrentVersion\Windows"
		strValueName = "Device"
		objRegistry.GetStringValue HKEY_CURRENT_USER,strKeyPath,strValueName,strDefaultPrinterValue
		strDefaultPrinterValue = UCase(Left(strDefaultPrinterValue,InStr(strDefaultPrinterValue,CHR(44))-1))
		'Check since line below is not acting as it should!!!
		If IsNull(strDefaultPrinterValue) Then
			DumpWindowsProfileInfo "WMI_GetStringValue","NOK"
		Else
			DumpWindowsProfileInfo "WMI_GetStringValue","OK"
		End If
		Set objRegistry = nothing
	End If
	GetDefaultPrinterFromRegistry = strDefaultPrinterValue
End Function
'----------------------------------------------------------------------------------------------------



Sub DumpWindowsProfileInfo (Level,Status)
	Const REG_SZ = 1
	Const REG_EXPAND_SZ = 2
	Const REG_BINARY = 3
	Const REG_DWORD = 4
	Const REG_MULTI_SZ = 7

	Dim ErrorNumber, ErrorDescription
	Dim StatusDir
	Dim ProfDebugLogDir, NewProfDebugDir, objProfDebugLog
	Dim colOperatingSystems, objOperatingSystem, objReg, WinVersion
	Dim objAccount
	Dim h
	Dim strKeyPath, arrValueNames(), arrValueTypes(), strValue, arrBytes(), strBytes, uByte, uValue
	Dim hh, BlankStream1, BlankStream2, RegInfoLength
	Dim UserProfiles, Profile
	
	ErrorNumber = Err.Number
	ErrorDescription = Replace(Err.Description,vbCRLF,vbNullString)
	Err.Clear
	On Error Goto 0
	Select Case Status
		Case "OK"
			StatusDir = "Good"
		Case "NOK"
			StatusDir = "Bad"
	End Select
	If objFSO.FolderExists(PrintQMigrationLogs & "Profile_Debug\" & StatusDir) Then
		ProfDebugLogDir = PrintQMigrationLogs & "Profile_Debug\" & StatusDir
	Else
		On Error Resume Next

		If Not objFSO.FolderExists(PrintQMigrationLogs & "Profile_Debug") Then
			Set NewProfDebugDir = objFSO.CreateFolder(PrintQMigrationLogs & "Profile_Debug")
		End If

		Set NewProfDebugDir = objFSO.CreateFolder(PrintQMigrationLogs & "Profile_Debug\" & StatusDir)
		If Err.Number Then
			Err.Clear
			Set NewProfDebugDir = nothing
			objOutput.Close
			Set objOutput = nothing
			Set objShell = nothing
			Set objFSO = nothing
			Set objWMIService = nothing
			Set objNetwork = nothing
			Set objRecordSet = nothing
			Set objCommand = nothing
			objConnection.Close
			Set objConnection = nothing
			WScript.Quit
		Else
			ProfDebugLogDir = PrintQMigrationLogs & "Profile_Debug\" & StatusDir
		End If
		On Error Goto 0
		Set NewProfDebugDir = nothing
	End If
	On Error Resume Next
	Set objProfDebugLog = objFSO.OpenTextFile(ProfDebugLogDir & CHR(92) & "ProfDebug_" & strUserDept & CHR(95) & lCase(strUserName) & CHR(95) & strComputerName & CHR(95) & UserLogTimeStamp & ".log", ForWriting, True)
	If Err.Number Then
		Err.Clear
		Set objProfDebugLog = nothing
		Set NewProfDebugDir = nothing
		objOutput.Close
		Set objOutput = nothing
		Set objShell = nothing
		Set objFSO = nothing
		Set objWMIService = nothing
		Set objNetwork = nothing
		Set objRecordSet = nothing
		Set objCommand = nothing
		objConnection.Close
		Set objConnection = nothing
		WScript.Quit
	End If
	On Error Goto 0
	
	If Level = "Shell_RegRead" Then
		objProfDebugLog.writeline "Log created at the first attempt to read the Registry using the Shell RegRead method"
		objProfDebugLog.writeline
	End If
	
	If Level = "WMI_GetStringValue" Then
		objProfDebugLog.writeline "Log created at the second attempt to read the Registry using the WMI GetStringValue method"
		objProfDebugLog.writeline
	End If
	
	'Corrupted profile would have Err.Number = -2147024894 and Err.Description would contain "Invalid root in registry key"
	If Status = "NOK" Then
		objProfDebugLog.writeline "The error number was >" & ErrorNumber & "< with description >" & ErrorDescription & "<"
	End If
	
	Set colOperatingSystems = objWMIService.ExecQuery("Select * from Win32_OperatingSystem")
	For Each objOperatingSystem In colOperatingSystems
		WinVersion = Left(objOperatingSystem.Version,3)
	Next
	objProfDebugLog.writeline "The Windows version is >" & WinVersion & "<"
	objProfDebugLog.writeline
	objProfDebugLog.writeline
	
	Set objReg = GetObject("winmgmts:!root/default:StdRegProv")
	Set objAccount = objWMIService.Get("Win32_UserAccount.Name='" & strUserName & "',Domain='" & Left(strDomain,InStr(strDomain,CHR(46))-1) & "'")
	strKeyPath = "SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\" & objAccount.SID

	objProfDebugLog.writeline "Profile information contained in the Registry:"
	objProfDebugLog.writeline "----------------------------------------------"
	objReg.EnumValues HKEY_LOCAL_MACHINE, strKeyPath, arrValueNames, arrValueTypes
	For h = 0 To UBound(arrValueNames)
		Select Case arrValueTypes(h)
			Case REG_SZ
				objReg.GetStringValue HKEY_LOCAL_MACHINE, strKeyPath, arrValueNames(h), strValue
				BlankStream1 = vbNullString
				BlankStream2 = vbNullString
				RegInfoLength = Len(arrValueNames(h))
				For hh = RegInfoLength to 20
					BlankStream1 = BlankStream1 & CHR(32)
				Next
				RegInfoLength = Len(arrValueNames(h) & BlankStream1 & "(REG_SZ)")
				For hh = RegInfoLength to 36
					BlankStream2 = BlankStream2 & CHR(32)
				Next
				objProfDebugLog.writeline arrValueNames(h) & BlankStream1 & "(REG_SZ)" & BlankStream2 & "=  " & strValue
			Case REG_EXPAND_SZ
				objReg.GetExpandedStringValue HKEY_LOCAL_MACHINE, strKeyPath, arrValueNames(h), strValue
				BlankStream1 = vbNullString
				BlankStream2 = vbNullString
				RegInfoLength = Len(arrValueNames(h))
				For hh = RegInfoLength to 20
					BlankStream1 = BlankStream1 & CHR(32)
				Next
				RegInfoLength = Len(arrValueNames(h) & BlankStream1 & "(REG_EXPAND_SZ)")
				For hh = RegInfoLength to 36
					BlankStream2 = BlankStream2 & CHR(32)
				Next
				objProfDebugLog.writeline arrValueNames(h) & BlankStream1 & "(REG_EXPAND_SZ)" & BlankStream2 & "=  " & strValue
			Case REG_BINARY
				objReg.GetBinaryValue HKEY_LOCAL_MACHINE, strKeyPath, arrValueNames(h), arrBytes
				strBytes = vbNullString
				For Each uByte in arrBytes
					strBytes = strBytes & Right("0" & Hex(uByte),2) & " "
				Next
				BlankStream1 = vbNullString
				BlankStream2 = vbNullString
				RegInfoLength = Len(arrValueNames(h))
				For hh = RegInfoLength to 20
					BlankStream1 = BlankStream1 & CHR(32)
				Next
				RegInfoLength = Len(arrValueNames(h) & BlankStream1 & "(REG_BINARY)")
				For hh = RegInfoLength to 36
					BlankStream2 = BlankStream2 & CHR(32)
				Next
				objProfDebugLog.writeline arrValueNames(h) & BlankStream1 & "(REG_BINARY)" & BlankStream2 & "=  " & strBytes
			Case REG_DWORD
				objReg.GetDWORDValue HKEY_LOCAL_MACHINE, strKeyPath, arrValueNames(h), uValue
				BlankStream1 = vbNullString
				BlankStream2 = vbNullString
				RegInfoLength = Len(arrValueNames(h))
				For hh = RegInfoLength to 20
					BlankStream1 = BlankStream1 & CHR(32)
				Next
				RegInfoLength = Len(arrValueNames(h) & BlankStream1 & "(REG_DWORD)")
				For hh = RegInfoLength to 36
					BlankStream2 = BlankStream2 & CHR(32)
				Next
				objProfDebugLog.writeline arrValueNames(h) & BlankStream1 & "(REG_DWORD)" & BlankStream2 & "=  " & CStr(uValue)
			Case REG_MULTI_SZ
				objReg.GetMultiStringValue hDefKey, strKeyPath, arrValueNames(h), arrValues				  				
				BlankStream1 = vbNullString
				BlankStream2 = vbNullString
				RegInfoLength = Len(arrValueNames(h))
				For hh = RegInfoLength to 20
					BlankStream1 = BlankStream1 & CHR(32)
				Next
				RegInfoLength = Len(arrValueNames(h) & BlankStream1 & "(REG_MULTI_SZ)")
				For hh = RegInfoLength to 36
					BlankStream2 = BlankStream2 & CHR(32)
				Next
				objProfDebugLog.writeline arrValueNames(h) & BlankStream1 & "(REG_MULTI_SZ)" & BlankStream2 & "=  "
				For Each strValue in arrValues
					objProfDebugLog.write "    " & strValue 
				Next
		End Select
	Next
	
	If Left(WinVersion,1) > 5 Then
		objProfDebugLog.writeline
		objProfDebugLog.writeline
		objProfDebugLog.writeline "Profile information queried via WMI:"
		objProfDebugLog.writeline "------------------------------------"
		Set UserProfiles = objWMIService.ExecQuery("select * from Win32_userprofile where SID = '" & objAccount.SID & "'")
		For Each Profile in UserProfiles
			Select Case Profile.Status
				Case 0
					objProfDebugLog.writeline "Status 			=  " & Profile.Status & " (Temporary)"
				Case 1
					objProfDebugLog.writeline "Status 			=  " & Profile.Status & " (Roaming)"
				Case 2
					objProfDebugLog.writeline "Status 			=  " & Profile.Status & " (Mandatory)"
				Case 3
					objProfDebugLog.writeline "Status			=  " & Profile.Status & " (Corrupted)"
			End Select
			objProfDebugLog.writeline "LastUseTime		=  " & Profile.LastUseTime
			objProfDebugLog.writeline "LastDownloadTime	=  " & Profile.LastDownloadTime
			objProfDebugLog.writeline "LastUploadTime		=  " & Profile.LastUploadTime
			objProfDebugLog.writeline "LocalPath		=  " & Profile.LocalPath
			objProfDebugLog.writeline "refCount		=  " & Profile.refCount
			objProfDebugLog.writeline "RoamingConfigured	=  " & Profile.RoamingConfigured
			objProfDebugLog.writeline "RoamingPath		=  " & Profile.RoamingPath
			objProfDebugLog.writeline "RoamingPreference	=  " & Profile.RoamingPreference
			objProfDebugLog.writeline "SID			=  " & Profile.SID
			objProfDebugLog.writeline "Special			=  " & Profile.Special
		Next
	End If
	
	objProfDebugLog.Close
	Set objProfDebugLog = nothing
End Sub
'----------------------------------------------------------------------------------------------------



Sub DumpSyntax
	objOutput.writeline
	objOutput.writeline "The Script syntax is as follows when using arguments :"
	objOutput.writeline
	objOutput.writeline vbTab & "[AnyFileName.csv]" & vbTab & vbTab & vbTab & "to run the script against mapping file [AnyFileName.csv] for all users calling the script"
	objOutput.writeline
	objOutput.writeline vbTab & "/CheckGroupMembership" & vbTab & vbTab & vbTab & "to only run the script against mapping file [" & strUserDept & "]-PrintQMig.csv for users members of AD group " & strUserDept & "-PrinterMigration"
	objOutput.writeline vbTab & "/CheckGroupMembership:[GroupName]" & vbTab & "to only run the script against mapping file [" & strUserDept & "]-PrintQMig.csv for users members of AD group [GroupName]"
	objOutput.writeline
	objOutput.writeline vbTab & "/Affix:[\\PrinterServer\PrintQueue]" & vbTab & "to add printer [\\PrinterServer\PrintQueue] for all users calling the script"
	objOutput.writeline
	objOutput.writeline vbTab & "/Remove:[\\PrinterServer\PrintQueue]" & vbTab & "to remove printer [\\PrinterServer\PrintQueue] for all users calling the script and having printer [\\PrinterServer\PrintQueue] installed"
	objOutput.writeline
	objOutput.writeline "Additionnally, the /Affix: or the /Remove: switches can be used in conjunction with either of the two possible /CheckGroupMembership switches for greater control"
End Sub
'----------------------------------------------------------------------------------------------------



Sub AffixPrinterToAffix (PrinterToAffix)
	PrinterToAffix = Replace(PrinterToAffix,CHR(46) & strDomain,vbNullString)
	If Len(Replace(PrinterToAffix,CHR(92),vbNullString)) = Len(PrinterToAffix) - 3 Then
		If InStr(InstalledPrintersList,PrinterToAffix) Then
			objOutput.writeline "Printer " & PrinterToAffix & " to affix is already present!"
			SingleTaskStatus = "<NOT_REQUIRED>"
		Else
			objOutput.writeline "Identified missing printer " & PrinterToAffix & " to install"
			If AddPrinter(PrinterToAffix) Then
				SingleTaskStatus = "Single printer affix task performed with success"
			Else
				SingleTaskStatus = "<FAILED>"
			End If
		End If
	Else
		objOutput.writeline "Error : " & PrinterToAffix & " to affix is not a valid printer name!"
		SingleTaskStatus = "<NOT_PERFORMED>"
	End If
End Sub
'----------------------------------------------------------------------------------------------------



Sub RemovePrinterToRemove (PrinterToRemove)
	Dim MatchFound : MatchFound = False

	PrinterToRemove = Replace(PrinterToRemove,CHR(46) & strDomain,vbNullString)
	If Len(Replace(PrinterToRemove,CHR(92),vbNullString)) = Len(PrinterToRemove) - 3 Then
		'The InstalledPrintersList mechanism used above in the Affix scenario cannot be used in this case since the native name of the printer to remove might be an FQDN. Therefore requiring the netPrinters array to be parsed
		If netPrinters(0) <> "No printers found" and netPrinters(0) <> "Cannot enumerate printers" Then
			For j = 0 to UBound(netPrinters)
				'Check if current printer is the default printer
				strCurrentPrinterShortUNC = Replace(UCase(netPrinters(j)),",DEFAULT",vbNullString)
				strCurrentPrinterShortUNC = Replace(strCurrentPrinterShortUNC,CHR(46) & strDomain,vbNullString)
				If strCurrentPrinterShortUNC = PrinterToRemove Then
					objOutput.writeline "Identified installed printer " & PrinterToRemove & " to remove"
					If InStr(netPrinters(j),",DEFAULT") Then
						objOutput.writeline "	Printer to remove " & PrinterToRemove & " is the current default printer and will not be removed until another default printer is set"
						SingleTaskStatus = "<NOT_PERFORMED>"
					Else
						If RemovePrinter(netPrinters(j)) Then
							SingleTaskStatus = "Single printer removal task performed with success"
						Else
							SingleTaskStatus = "<FAILED>"
						End If
					End If
					MatchFound = True
				End If
			Next
		End If
		If MatchFound = False Then
			objOutput.writeline "Printer " & PrinterToRemove & " to remove not present"
			SingleTaskStatus = "<NOT_REQUIRED>"
		End If
	Else
		objOutput.writeline "Error : " & PrinterToRemove & " to remove is not a valid printer name!"
	End If
End Sub
'----------------------------------------------------------------------------------------------------



Sub MigratePrinters
	Dim n, o
	Dim MatchingMappingEntry, MappingEntry, ObsoletePrintServer
	Dim arrAddedPrintersList
	Dim DefaultPrinter, DefaultPrinterName, strUserPrinterUNC
	Dim arrCurrentPrinterUNC, strPrintServerCurrentPrinter
	
	For j = 0 to UBound(netPrinters)
		MatchingMappingEntry = "<NOT_FOUND>"
		netPrinters(j) = Mid(UCase(netPrinters(j)),5)
		'Check if current printer is the default printer
		If InStr(netPrinters(j),",DEFAULT") Then
			DefaultPrinter = True
			'Truncating, removing word 'DEFAULT' from Array
			strCurrentPrinterShortUNC =	Replace(UCase(netPrinters(j)),",DEFAULT",vbNullString)
		Else
			DefaultPrinter = False
			strCurrentPrinterShortUNC = UCase(netPrinters(j))
		End If
		strCurrentPrinterShortUNC = Replace(strCurrentPrinterShortUNC,CHR(46) & strDomain,vbNullString)
		If InStr(PrintersToAffix,strCurrentPrinterShortUNC) Then
			If InStr(PrintersToAffix,CHR(124) & strCurrentPrinterShortUNC) Then
				PrintersToAffix = Replace(PrintersToAffix,CHR(124) & strCurrentPrinterShortUNC,vbNullString)
			Else
				If InStr(PrintersToAffix,strCurrentPrinterShortUNC & CHR(124)) Then
					PrintersToAffix = Replace(PrintersToAffix,strCurrentPrinterShortUNC & CHR(124),vbNullString)
				Else
					PrintersToAffix = Replace(PrintersToAffix,strCurrentPrinterShortUNC,vbNullString)
				End If
			End If
		End If
		For Each MappingEntry in AllMappings
			If MappingEntry <> vbNullString and InStr(MappingEntry,CHR(44)) and Left(MappingEntry,2) = "\\" Then
				MappingEntry = Replace(UCase(Replace(MappingEntry,CHR(13),vbNullString)),CHR(46) & strDomain,vbNullString)
				If InStr(MappingEntry,strCurrentPrinterShortUNC) Then
					If Len(Replace(MappingEntry,CHR(92),vbNullString)) = Len(MappingEntry)-6 Then
						If strCurrentPrinterShortUNC = Mid(MappingEntry,1,Instr(MappingEntry,CHR(44))-1) Then
							MatchingMappingEntry = MappingEntry
							Exit For
						Else
							If strCurrentPrinterShortUNC = Mid(MappingEntry,Instr(MappingEntry,CHR(44))+1) Then
								MatchingMappingEntry = "<ALREADY_MIGRATED>"
								Exit For
							End If
						End If
					Else
						'Department compulsory Printer case
						If Mid(MappingEntry,Instr(MappingEntry,CHR(44))+1) = "INSTALL" Then
							MatchingMappingEntry = "<DEPT_WIDE_PRINTER>"
							Exit For
						End If
						'Obsolete Printer case
						If Len(Replace(Mid(MappingEntry,1,Instr(MappingEntry,CHR(44))-1),CHR(92),vbNullString)) = Len(Mid(MappingEntry,1,Instr(MappingEntry,CHR(44))-1))-3 and Mid(MappingEntry,Instr(MappingEntry,CHR(44))+1) = "DELETE" Then
							If Mid(MappingEntry,1,Instr(MappingEntry,CHR(44))-1) = strCurrentPrinterShortUNC Then
								MatchingMappingEntry = "<OBSOLETE_PRINTER>"
								Exit For
							End If
						End If
					End If
				Else
					If Len(Replace(MappingEntry,CHR(92),vbNullString)) = Len(MappingEntry)-2 Then
						'Obsolete Server case
						If Len(Replace(Mid(MappingEntry,1,Instr(MappingEntry,CHR(44))-1),CHR(92),vbNullString)) = Len(Mid(MappingEntry,1,Instr(MappingEntry,CHR(44))-1))-2 and Mid(MappingEntry,Instr(MappingEntry,CHR(44))+1) = "DELETE" Then
							ObsoletePrintServer = Mid(MappingEntry,1,Instr(MappingEntry,CHR(44))-1)
							If InStr(strCurrentPrinterShortUNC,ObsoletePrintServer) Then
								MatchingMappingEntry = "<ON_OBSOLETE_SERVER>"
								Exit For
							End If
						End If
					End If
				End If
			End If
		Next

		If MatchingMappingEntry <> "<NOT_FOUND>" and MatchingMappingEntry <> "<ALREADY_MIGRATED>" and MatchingMappingEntry <> "<DEPT_WIDE_PRINTER>" and MatchingMappingEntry <> "<OBSOLETE_PRINTER>" and MatchingMappingEntry <> "<ON_OBSOLETE_SERVER>" Then
			arrOldToNewPrinter = split(MatchingMappingEntry,CHR(44))
			strOldPrinterUNC = arrOldToNewPrinter(0) 'Check but it seems it is not really needed!
			strNewPrinterUNC = arrOldToNewPrinter(1)
			objOutput.writeline "Identified printer " & strCurrentPrinterShortUNC & " to migrate to " & strNewPrinterUNC
			If InStr(InstalledPrintersList,strNewPrinterUNC) = 0 and InStr(AddedPrintersList,strNewPrinterUNC) = 0 Then
				EnrolledPrinters = EnrolledPrinters + 1
				If AddPrinter(strNewPrinterUNC) Then
					If DefaultPrinter Then 
						SetDefaultPrinter (strNewPrinterUNC)
						arrAddedPrintersList = split(AddedPrintersList,CHR(124))
						For n = 0 to UBound(arrAddedPrintersList)
							If arrAddedPrintersList(n) = strNewPrinterUNC Then
								arrAddedPrintersList(n) = arrAddedPrintersList(n) & CHR(44) & "DEFAULT"
							End If
						Next
						AddedPrintersList = Join (arrAddedPrintersList,CHR(124))
					End If
					If Not RemovePrinter(netPrinters(j)) Then
						UnremovablePrinters = UnremovablePrinters + 1
					End If
					MigratedPrinters = MigratedPrinters + 1
				Else 
					objOutput.writeline "	Error: Printer " & strCurrentPrinterShortUNC & " not migrated!"
				End If
			Else
				If InStr(InstalledPrintersList,strNewPrinterUNC) Then
					objOutput.writeline "	Printer " & strCurrentPrinterShortUNC & " is to be replaced by " & strNewPrinterUNC & " which is already installed"
					If DefaultPrinter Then
						objOutput.writeline "	Printer " & strCurrentPrinterShortUNC & " is currently the default printer. Setting already installed " & strNewPrinterUNC & " as new default"
						'Using native default printer name from netPrinters array instead of strNewPrinterUNC, just in case the print server is defined with an FQDN rather than a NetBIOS name
						For o = 0 to UBound(netPrinters)
							If Instr(UCase(netPrinters(o)),strDomain) Then
								strUserPrinterUNC = Mid(netPrinters(o),1,InStr(netPrinters(o),CHR(46))-1) & Replace(Mid(UCase(netPrinters(o)),InStr(netPrinters(o),CHR(46))+1),strDomain,vbNullString)
							Else
								strUserPrinterUNC = netPrinters(o)
							End If
							If Replace(UCase(strUserPrinterUNC),",DEFAULT",vbNullString) = strNewPrinterUNC Then
								DefaultPrinterName = Replace(netPrinters(o),",DEFAULT",vbNullString)
								Exit For
							End If
						Next
						SetDefaultPrinter(DefaultPrinterName)
					End If
				End If
				If InStr(AddedPrintersList,strNewPrinterUNC) Then
					objOutput.writeline "	Printer " & strCurrentPrinterShortUNC & " is to be replaced by " & strNewPrinterUNC & " which has just been added"
					If DefaultPrinter and InStr(AddedPrintersList,"DEFAULT") = 0 Then 
						objOutput.writeline "	Printer " & strCurrentPrinterShortUNC & " is currently the default printer. Setting just added " & strNewPrinterUNC & " as new default"
						SetDefaultPrinter(strNewPrinterUNC)
					End If
				End If
				objOutput.writeline "	Skipping duplicate installation of " & strNewPrinterUNC & " and removing " & strCurrentPrinterShortUNC
				If RemovePrinter(netPrinters(j)) Then
					RemovedPrinters = RemovedPrinters + 1
				Else
					UnremovablePrinters = UnremovablePrinters + 1
				End If
			End If
		Else
			If MatchingMappingEntry = "<ALREADY_MIGRATED>" Then
				objOutput.writeline "Printer " & strCurrentPrinterShortUNC & " has already been migrated"
			End If
			If MatchingMappingEntry = "<DEPT_WIDE_PRINTER>" Then
				objOutput.writeline "Printer " & strCurrentPrinterShortUNC & " is an already present " & strUserDept & " Department wide printer"
			End If
			If MatchingMappingEntry = "<OBSOLETE_PRINTER>" Then
				objOutput.writeline "Identified printer " & strCurrentPrinterShortUNC & " to remove"
				If Not DefaultPrinter Then 
					If RemovePrinter(netPrinters(j)) Then
						RemovedPrinters = RemovedPrinters + 1
					Else
						UnremovablePrinters = UnremovablePrinters + 1
					End If
				Else
					objOutput.writeline "	Obsolete printer " & strCurrentPrinterShortUNC & " is the current default printer and will not be removed until another default printer is set"
				End If
			End If
			If MatchingMappingEntry = "<ON_OBSOLETE_SERVER>" Then
				objOutput.writeline "Identified printer " & strCurrentPrinterShortUNC & " on obsolete server " & ObsoletePrintServer & " to remove"
				If Not DefaultPrinter Then 
					If RemovePrinter(netPrinters(j)) Then
						RemovedPrinters = RemovedPrinters + 1
					Else
						UnremovablePrinters = UnremovablePrinters + 1
					End If
				Else
					objOutput.writeline "	Printer " & strCurrentPrinterShortUNC & " on obsolete server " & ObsoletePrintServer & " is the current default printer and will not be removed until another default printer is set"
				End If
			End If
			If MatchingMappingEntry = "<NOT_FOUND>" Then
				If Len(Replace(strCurrentPrinterShortUNC,CHR(92),vbNullString)) = Len(strCurrentPrinterShortUNC)-3 Then
					objOutput.writeline "Printer " & strCurrentPrinterShortUNC & " not found in mapping list"
				Else
					objOutput.writeline "Printer " & strCurrentPrinterShortUNC & " is not a valid network printer name and should be handled manually"
				End If
			End If
		End If
	Next
End Sub
'----------------------------------------------------------------------------------------------------



Function RemovePrinter (OldPrinter)
	Const boolForce = True
	Const boolUpdateProfile = True
	Dim OldPrinterShortUNC
	Dim RunDLLCommandLine
	Dim boolPrinterRemoved : boolPrinterRemoved = True
	Dim objReg
	Dim p
	Dim arrMappedPrinters, arrValueTypes

	OldPrinter = Replace(OldPrinter,",DEFAULT",vbNullString)
	OldPrinterShortUNC = Replace(UCase(OldPrinter),CHR(46) & strDomain,vbNullString)
	On Error Resume Next
	objNetwork.RemovePrinterConnection OldPrinter, boolForce, boolUpdateProfile
	If Err.Number = -2147022646 Then
		objOutput.writeline "	Warning: Obsolete printer " & OldPrinterShortUNC & " could not be removed using VBScript Network object RemovePrinterConnection method [" & Replace(Err.Description,vbCRLF,vbNullString) & CHR(93)
		Err.Clear
		RunDLLCommandLine = "RUNDLL32 PRINTUI.DLL, PrintUIEntry /dn /n " & OldPrinter
		objShell.Run RunDLLCommandLine, 0, True
		WScript.Sleep 100
		Set objReg = GetObject("winmgmts:!root/default:StdRegProv")
		objReg.EnumValues HKEY_CURRENT_USER, "Software\Microsoft\Windows NT\CurrentVersion\Devices", arrMappedPrinters, arrValueTypes
		For p = 0 To UBound(arrMappedPrinters)
			If UCase(arrMappedPrinters(p)) = UCase(OldPrinter) or InStr(UCase(arrMappedPrinters(p)),UCase(OldPrinter)) Then
				boolPrinterRemoved = False
				Exit For
			End If
		Next		
		If boolPrinterRemoved = True Then
			objOutput.writeline "	Obsolete printer " & OldPrinterShortUNC & " successfully removed using PrintUI"
		Else
			objOutput.writeline "	Error: Obsolete printer " & OldPrinterShortUNC & " could not be removed using PrintUI"
		End If 
		Set objReg = nothing
	Else
		objOutput.writeline "	Obsolete printer " & OldPrinterShortUNC & " successfully removed using VBScript RemovePrinterConnection method"
	End If
	On Error GoTo 0
	RemovePrinter = boolPrinterRemoved
End Function
'----------------------------------------------------------------------------------------------------



Function CheckPrinterToAdd (NewPrinter)
	Dim arrNewPrinterUNC, Server, Printer, PrinterCN
	Dim objExec
	Dim NetViewLine, SharedPrinter, FailedNetView
	Dim boolOKToInstall : boolOKToInstall = False
	Dim strBase
	Dim strLDAPQuery
	Dim strPrinterADUNCName, strPrinterDN, arrPrinterDN, strPrinterADPath, strPrinterADName, strADShortPrintServerName
	Dim q
	
	If Len(Replace(NewPrinter,CHR(92),vbNullString)) = Len(NewPrinter)-3 Then 
		arrNewPrinterUNC = split(NewPrinter,CHR(92))
		Server = arrNewPrinterUNC(2)
		Printer = arrNewPrinterUNC(3)
		PrinterCN = Server & "-" & Printer
		
		'The Net View check against the target server was added as a precautionary measure to ensure the new Print Queues 
		'defined in the mapping file effectively exist as corresponding shared Print Queues on the target Print Server.
		'First, using the NET VIEW command, check if a shared printer exists with that Printer name on the new Print Server
		Set objExec = objShell.Exec("cmd /c net view \\" & Server & " |find ""Print"" /i") 
		Do 
			NetViewLine = objExec.StdOut.ReadLine
			If InStr(NetViewLine,CHR(32)) Then
				SharedPrinter = Mid(NetViewLine,1,InStr(NetViewLine,CHR(32))-1)
			End If
		Loop Until UCase(SharedPrinter) = Printer or NetViewLine = vbNullString
		
		If NetViewLine <> vbNullString Then
			'Look first if the New Printer to add is defined in the appropriate AD OU containing printers 
			strBase = PrintServersOU & strDomainDN & CHR(62)  

			'LDAP Query to return the UNC Name of the New Printer assuming the Share Name in the mapping table is equal to the printer's AD CN (Common Name)
			strLDAPQuery = strBase & CHR(59) & "(&(objectCategory=printQueue)(cn=" & PrinterCN & "))" & ";uNCName;subtree"
			objCommand.CommandText = strLDAPQuery
			Set objRecordSet = objCommand.Execute
			If Not objRecordset.EOF Then
				strPrinterADUNCName = Replace(UCase(objRecordSet.Fields("uNCName").Value),CHR(46) & strDomain,vbNullString)
				If  strPrinterADUNCName = Ucase(NewPrinter) Then
					boolOKToInstall = True
				Else
					objOutput.writeline "	Error: Target printer " & NewPrinter & " is not shared with this name!"
					objOutput.writeline "	Error: Target printer " & NewPrinter & " is in fact shared as " & strPrinterADUNCName
				End If
			Else
				'Could not locate Printer in appropriate OU, look for this printer in entire AD scope
				strBase = "<LDAP://" & strDomainDN & CHR(62)
				strLDAPQuery = strBase & CHR(59) & "(&(objectCategory=printQueue)(cn=" & PrinterCN & "));distinguishedName;subtree"
				objCommand.CommandText = strLDAPQuery
				Set objRecordSet = objCommand.Execute
				If Not objRecordset.EOF Then
					strPrinterDN = objRecordSet.Fields("distinguishedName").Value
					arrPrinterDN = split(strPrinterDN,CHR(44))
					strPrinterADPath = vbNullString
					For q = UBound(arrPrinterDN) to 0 Step -1
						If InStr(arrPrinterDN(q),"OU=") Then
							strPrinterADPath = strPrinterADPath & CHR(47) & arrPrinterDN(q)
						End If
					Next
					If Instr(strPrinterDN,PrintServersOU) = 0 Then
						objOutput.writeline "	Error: Target printer shared as " & NewPrinter & " is not published in the right AD OU container!"
						objOutput.writeline "	Error: Target printer shared as " & NewPrinter & " is wrongfully published in " & strPrinterADPath
					End If
				Else
					strLDAPQuery = strBase & CHR(59) & "(&(objectCategory=printQueue)(printerName=" & Printer & ")(shortServerName=" & Server & "));name;subtree"
					objCommand.CommandText = strLDAPQuery
					Set objRecordSet = objCommand.Execute
					If Not objRecordset.EOF Then
						strPrinterADName = objRecordSet.Fields("name").Value
						objOutput.writeline "	Error: Target printer " & NewPrinter & " is published in AD under a different printer name!"
						objOutput.writeline "	Error: Target printer " & NewPrinter & " is in fact published in AD as " & strPrinterADName
						If InStr(NewPrinter,"S-PDFCODE-") Then
							objOutput.writeline "	Installing PDF CODE printer despite this AD inconsistency (Exception tolerated)"
							boolOKToInstall = True
						End If
					Else
						objOutput.writeline "	Error: Target printer " & NewPrinter & " IS NOT published in AD at all!"
					End If
				End If
			End If
		Else
			FailedNetView = "<SERVER_INACCESSIBLE>"
			Set objExec = objShell.Exec("cmd /c net view \\" & Server)
			If inStr(objExec.StdOut.ReadLine,"Shared resources at") Then
				FailedNetView = "<NO_SHARED_PRINTER_ON_SERVER>"
			End If
			If MappingFile <> vbNullString Then
				If FailedNetView = "<NO_SHARED_PRINTER_ON_SERVER>" Then
					objOutput.writeline "	Error: Target printer " & Printer & " specified in printer mapping file " & MappingFile & " does not exist on server " & Server & "!"
				Else
					objOutput.writeline "	Error: Target Server " & Server & " specified in printer mapping file " & MappingFile & " cannot be contacted!"
				End If
			Else
				If FailedNetView = "<NO_SHARED_PRINTER_ON_SERVER>" Then
					objOutput.writeline "	Error: Target printer " & Printer & " specified in the script command line does not exist on server " & Server & "!"
				Else
					objOutput.writeline "	Error: Target Server " & Server & " specified in the script command line cannot be contacted!"
				End If
			End If
		End If
	Else
		If MappingFile <> vbNullString Then
			objOutput.writeline "	Error: Target printer " & NewPrinter & " specified in printer mapping file " & MappingFile & " is not a valid UNC name!"
		Else
			objOutput.writeline "	Error: Target printer " & NewPrinter & " specified in the script command line is not a valid UNC name!"
		End If
	End If
	
	If boolOKToInstall = False Then
		RejectedPrinters = RejectedPrinters + 1
	End If
	
	Set objExec = nothing

 	CheckPrinterToAdd = boolOKToInstall
End Function
'----------------------------------------------------------------------------------------------------



Function AddPrinter (NewPrinter)
	If InStr(NewPrinter,"S-PDFCODE-") or CheckPrinterToAdd(NewPrinter) Then
		On Error Resume Next
		objNetwork.AddWindowsPrinterConnection(NewPrinter)
		If Err.Number = -2147023099 Then
			Err.Clear
			On Error Goto 0
			objOutput.writeline "	Error: " & OSVersion & " driver for printer " & Mid(NewPrinter,InStrRev(NewPrinter,CHR(92))+1) & " to install on server " & Mid(NewPrinter,3,InStrRev(NewPrinter,CHR(92))-3) & " is not loaded!"
			RejectedPrinters = RejectedPrinters + 1
			AddPrinter = False
		Else
			On Error Goto 0
			If AddedPrintersList = vbNullString Then
				AddedPrintersList = NewPrinter
			Else
				AddedPrintersList = AddedPrintersList & CHR(124) & NewPrinter
			End If
			objOutput.writeline "	Added new printer " & NewPrinter
			AddPrinter = True
		End If
	Else
		AddPrinter = False
	End If
End Function
'----------------------------------------------------------------------------------------------------



Sub SetDefaultPrinter (NewDefaultPrinter)
	Dim RunDLLCommandLine
	Dim DefaultPrinterFromRegistry : DefaultPrinterFromRegistry = vbNullString

	On Error Resume Next
    objNetwork.SetDefaultPrinter(NewDefaultPrinter)
	If Err.Number Then
		Err.Clear
		On Error Goto 0
		objOutput.writeline "	Warning: Newly installed target printer " & NewDefaultPrinter & " could not be set as default printer using VBScript Network object SetDefaultPrinter method [" & Err.Number & CHR(44) & Replace(Err.Description,vbCRLF,vbNullString) & CHR(93)
		RunDLLCommandLine = "RUNDLL32 PRINTUI.DLL, PrintUIEntry /y /n " & NewDefaultPrinter
		objShell.Run RunDLLCommandLine, 0, True
		DefaultPrinterFromRegistry = objShell.RegRead("HKCU\Software\Microsoft\Windows NT\CurrentVersion\Windows\Device")
		DefaultPrinterFromRegistry = UCase(Left(DefaultPrinterFromRegistry,InStr(DefaultPrinterFromRegistry,CHR(44))-1))
		If DefaultPrinterFromRegistry = NewDefaultPrinter Then
			objOutput.writeline "	Printer " & NewDefaultPrinter & " successfully set as default printer using PrintUI"
		Else
			objOutput.writeline "	Error: Printer " & NewDefaultPrinter & " could not be set as default printer using PrintUI"
		End If
	Else
		On Error Goto 0
		objOutput.writeline "	Set default printer to " & NewDefaultPrinter
	End If
End Sub
'----------------------------------------------------------------------------------------------------



Function AffixCompulsoryPrinters (PrintersToAffix)
	Dim arrPrintersToAffix
	Dim r
	
	'Affixing Department wide compulsory printers
	AffixedCompulsoryPrinters = 0
	If InStr(PrintersToAffix,CHR(124)) Then
		arrPrintersToAffix = split(PrintersToAffix,CHR(124))
		For r = 0 to UBound(arrPrintersToAffix)
			If InStr(InstalledPrintersList,arrPrintersToAffix(r)) = 0 and InStr(AddedPrintersList,arrPrintersToAffix(r)) = 0 Then
				objOutput.writeline "Identified " & strUserDept & " Department wide printer " & arrPrintersToAffix(r) & " to install"
				If AddPrinter(arrPrintersToAffix(r)) Then
					AffixedCompulsoryPrinters = AffixedCompulsoryPrinters + 1
				End If
			End If
		Next
	Else
		If InStr(InstalledPrintersList,PrintersToAffix) = 0 and InStr(AddedPrintersList,PrintersToAffix) = 0 Then
			objOutput.writeline "Identified missing " & strUserDept & " Department wide printer " & PrintersToAffix & " to install"
			If AddPrinter(PrintersToAffix) Then
				AffixedCompulsoryPrinters =  AffixedCompulsoryPrinters + 1
			End If
		End If
	End If
End Function
'----------------------------------------------------------------------------------------------------



Sub MigrationSummary
	objOutput.writeline "==================================================================================================="
	objOutput.writeline
	If SinglePrinterTask Then
		objOutput.writeline "Printer Single Task summary for " &strUserDept & " user " & strUserName & " [" & strUserCN & "] :"
	Else
		If WipeAllPrintersTask Then
			objOutput.writeline "Printer Wipe Task summary for " &strUserDept & " user " & strUserName & " [" & strUserCN & "] :"
		Else
			objOutput.writeline "Printer Migration summary for " &strUserDept & " user " & strUserName & " [" & strUserCN & "] :"
		End If
	End If
	objOutput.writeline
	If (PrintMigGroupError = vbNullString or PrintMigGroupError = " recursively") and MappingFileError = vbNullString Then
		If EnrolledPrinters > 0 or PrintersToAffix <> vbNullString Then
			If MigratedPrinters > 0 Then
				objOutput.writeline "Printers migrated    : " & MigratedPrinters & CHR(47) & EnrolledPrinters
			End If
			If RejectedPrinters > 0 Then
				If PrintersToAffix = vbNullString Then
					objOutput.writeline "Printers rejected    : " & RejectedPrinters & CHR(47) & EnrolledPrinters
				Else
					objOutput.writeline "Printers rejected    : " & RejectedPrinters & CHR(47) & EnrolledPrinters + (Len(PrintersToAffix) - Len(Replace(PrintersToAffix,CHR(124),vbNullString)) + 1)
				End If
			End If
		Else
			If RemovedPrinters = 0 and UnremovablePrinters = 0 and AffixedPrinters = 0 Then
				If PrinterToAffix = vbNullString and PrinterToRemove = vbNullString Then
					objOutput.writeline "No action required"
				Else
					If InStr(SingleTaskStatus,"success") Then
						objOutput.writeline SingleTaskStatus
					Else
						If SingleTaskStatus = "<FAILED>" Then
							objOutput.writeline "Single printer task failed!"
						Else
							If SingleTaskStatus = "<NOT_REQUIRED>" Then
								objOutput.writeline "Single printer task not required!"
							Else
								If WipeAllPrintersTask Then
									objOutput.writeline "No printers to remove!"
								Else
									objOutput.writeline "Single printer task not performed!"
								End If
							End If
						End If
					End If
				End If
			End If
		End If
		If AffixedPrinters > 0 Then
			objOutput.writeline "Department Printers affixed  : " & AffixedPrinters
		End If
		If RemovedPrinters > 0 Then
			objOutput.writeline
			objOutput.writeline "Printers removed     : " & RemovedPrinters & CHR(47) & netPrintersCount
		End If
		If UnremovablePrinters > 0 Then
			objOutput.writeline "Unremovable printers : " & UnremovablePrinters
		End If
		If EnrolledPrinters > 0 and netPrintersCount - MigratedPrinters - RemovedPrinters > 0 Then
			If RemovedPrinters = 0 Then
				objOutput.writeline
			End If
			objOutput.writeline "Printers unchanged   : " & netPrintersCount - MigratedPrinters - RemovedPrinters & CHR(47) & netPrintersCount
		End If
	Else
		If PrintMigGroupError <> vbNullString and PrintMigGroupError = " recursively" and PrintMigGroupError <> "NoGroupCheck" Then
			objOutput.writeline PrintMigGroupError
			PrintMigGroupError = Null
		Else
			If MappingFileError = vbNullString Then
				objOutput.writeline "No action taken"
			End If
		End If
		If MappingFileError <> vbNullString Then
			objOutput.writeline MappingFileError
			MappingFileError = Null
		End If
	End If
	If MappingFileWarning <> vbNullString Then
		objOutput.writeline
		objOutput.writeline MappingFileWarning
	End If
    boolFinished = True
End Sub
'----------------------------------------------------------------------------------------------------



Function TimeTracker
	TimeTracker = Timer()
	If InStr(TimeTracker,CHR(46))<>0 Then
		TimeTracker = Mid(TimeTracker,1,InStr(TimeTracker,CHR(46))-1)
	End If
	If InStr(TimeTracker,CHR(44))<>0 Then
		TimeTracker = Mid(TimeTracker,1,InStr(TimeTracker,CHR(44))-1)
	End If
	TimeTracker = DatePart("y",Date) & CHR(124) & TimeTracker
End Function
'----------------------------------------------------------------------------------------------------



Function ComputeDuration (StartTrack,EndTrack)
	Dim arrStartTrack, arrEndTrack, StartDayOfYear, EndDayOfYear
	Dim TimerDelta, TimerHours, TimerMinutes, TimerSeconds
	
	arrStartTrack = split(StartTrack,CHR(124))
	arrEndTrack = split(EndTrack,CHR(124))
	StartDayOfYear = arrStartTrack(0)
	EndDayOfYear = arrEndTrack(0)
	
	If arrEndTrack(0) = arrStartTrack(0) Then
		ComputeDuration = "Duration : "
		TimerDelta = arrEndTrack(1) - arrStartTrack(1)
	Else
		If arrEndTrack(1) < arrStartTrack(1) Then
			If arrEndTrack(0) - arrStartTrack(0) = 1 Then
				ComputeDuration = "Duration : "
			Else
				If arrEndTrack(0) - arrStartTrack(0) = 2 Then
					ComputeDuration = "Duration : 1 Day, "
				Else
					ComputeDuration = "Duration : " & arrEndTrack(0) - arrStartTrack(0) - 1 & " Days, "
				End If
			End If
			TimerDelta = 86400 - arrStartTrack(1) + arrEndTrack(1)
		Else
			If arrEndTrack(0) - arrStartTrack(0) = 1 Then
				ComputeDuration = "Duration : 1 Day, "
			Else
				ComputeDuration = "Duration : " & arrEndTrack(0) - arrStartTrack(0) & " Days, "
			End If
			TimerDelta = arrEndTrack(1) - arrStartTrack(1)
		End If
	End If
	
	TimerHours = Int(TimerDelta / (60*60))
	TimerMinutes = Int((TimerDelta Mod (60*60)) / 60)
	TimerSeconds = (TimerDelta Mod (60*60)) Mod (60)
	
	Select Case TimerHours
		Case 0
			Select Case TimerMinutes
				Case 0
					If TimerSeconds = 0 or TimerSeconds = 1 Then
						ComputeDuration = ComputeDuration & TimerSeconds & " Second"
					Else
						ComputeDuration = ComputeDuration & TimerSeconds & " Seconds"
					End If
				Case 1
					ComputeDuration = ComputeDuration & TimerMinutes & " Minute : " & TimerSeconds & " Seconds"
				Case Else
					ComputeDuration = ComputeDuration & TimerMinutes & " Minutes : " & TimerSeconds & " Seconds"
			End Select
		Case 1
			ComputeDuration = ComputeDuration & TimerHours & " Hour : " & TimerMinutes & " Minutes : " & TimerSeconds & " Seconds"
		Case Else
			ComputeDuration = ComputeDuration & TimerHours & " Hours : " & TimerMinutes & " Minutes : " & TimerSeconds & " Seconds"
	End Select
End Function
'----------------------------------------------------------------------------------------------------



Sub MigrationClosure
	'Compute process duration and write into log before cleaning up and exiting
	TimeTrack(1) = TimeTracker()
	objOutput.writeline
	objOutput.writeline "___________________________________________________________________________________________________"
	objOutput.writeline
	objOutput.writeline ComputeDuration(TimeTrack(0),TimeTrack(1))
	
	If boolFinished Then 
		objOutput.Close
		Set objOutput = nothing
	End If
End Sub
'----------------------------------------------------------------------------------------------------



' ***************************************************************************************************
' ***************************************************************************************************