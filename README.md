# VBScript


Author       : Cary GARVIN  
Contact      : cary(at)garvin.tech  
LinkedIn     : https://www.linkedin.com/in/cary-garvin-99909582  
GitHub       : https://github.com/carygarvin/  


Script Name  : Migrate_PrintQs.vbs  
Version      : 1.0  
Release date : 07/02/2015 (CET)  
History      : The present script has been used by large organizations to successfully migrate tens of thousands of network printers from old to new Print Servers. A lot of safeguards have been buil into this Script.  
Purpose      : The present script is to be used in the scope of a Print Server migration whereby Network Print Queues are migrated from one old Print Server (to be decommissioned) to a new Print server.  
               The script will take care of remapping all of user Print Queues based on information contained in a mappings or correspondence file with each line in the format `'\\OldPrintServer\OldPrintQueueName,\\NewPrintServer\NewPrintQueueName'`.  
               On top of that the script has the ability to remove (Remove) some printers altogether or Add (Affix) one or more printers deemed compulsory.  

# Script information:
Script to migrate user printers from one Print Server to another based on a correspondence/mapping file holding Old Print Queue to New Print Queue mappings.  
The present Script is best invoked during the Login Script through Active Directory Group membership or interactively by specifying any specific mapping file to use as a parameter  
The present Script uses for maximum reliability several methods in order to identify user printers.  
The present Script has many features as follows:  
* Migrate user PrintQueues at logon or interactively based on information contained in the specified mappings file  
* Add one or more compulsory printers which all users not to have access to.  
* Unequivocally remove obsolete printers.  

As stated, the present script can be invoked either from a Logon Script in which the current user can have its group membership tested and if validated, call the script. By default, the mappings file that will be used will match the user's OU. This allows to have mappings file specific to each departmernt/bu in case of large organizations.  
Here's an example of how it can be called from within a "parent" VBScript Logon Script provided the Group's Distinguished Name is stored in strGroupDN and a binding has been made to the user object through objUser:  
  
                          Set objGroup = GetObject("LDAP://" & strGroupDN)
                          If objGroup.IsMember("LDAP://" & objUser.UserName) Then
                              objShell.Run "Migrate_PrintQs.vbs"
                          EndIf  
Alternatively, the script can be run from a Command Line (`cscript Migrate-PrintQs.vbs`).  
The script's Remove or Affix feature can be invoked either through Command Line switches when invoking the script (mostly used in interactive cases) or through specific fomratting of mappings with the mappings file (mostly used via a Logon Script).  

# Script usage:  
## Command Line switches:  
* FileName.csv  
* /Affix:  
* /Remove:  
* /RemoveAllPrinters  
* /CheckGroupMembership  
* /CheckGroupMembership:<CustomGroupName>  

## Command Line Examples:  
        Migrate_PrintQs.vbs PrintMigTable.csv                                   ==>      [This will migrate current Print Queues based on the information inside specified 'PrintMigTable.csv' file. This file is to be posted on the Network Share specified in the 'PrintQMappingsRepo' variable]  
        Migrate_PrintQs.vbs /Affix:\\ContosoNewPrtSrv\NewPrintQueueName			    ==>      [This will add a mapping to '\\ContosoNewPrtSrv\NewPrintQueueName' if none already exists]  
        Migrate_PrintQs.vbs /Remove:\\ContosoOldPrtSrv\OldPrintQueueName		    ==>      [This will remove any mapping to '\\ContosoOldPrtSrv\OldPrintQueueName' if any exists]  
        Migrate_PrintQs.vbs /RemoveAllPrinters                                  ==>      [This will remove all of user's printers]  
        Migrate_PrintQs.vbs /CheckGroupMembership                               ==>      [This will tell the script to act as if it is run within the Logon Script, meaning that the Mappings table to use is the default computed one for the user's devised Department.]  
        Migrate_PrintQs.vbs /CheckGroupMembership:PrintMigUsers                 ==>      [Same as above but for special cases where the user does not comply to the Department OU = Group prefix = Mappings CSV file prefix paradigm. The migration will take place based on the Mappings table from the user's Department OU]  

## Migration action Examples via mappings file (assuming Print Server migration from 'ContosoOldPrtSrv1' to 'ContosoNewPrtSrv1'):  
        \\ContosoOldPrtSrv1\OldPrtQ1,\\ContosoNewPrtSrv1\NewPrtQ1               ==>      Print Queue '\\ContosoOldPrtSrv1\OldPrtQ1' will be replaced by '\\ContosoNewPrtSrv1\NewPrtQ1'  
        \\ContosoOldPrtSrv1\OldPrtQ2,\\ContosoNewPrtSrv1\NewPrtQ2               ==>      Print Queue '\\ContosoOldPrtSrv1\OldPrtQ2' will be replaced by '\\ContosoNewPrtSrv1\NewPrtQ2'  
        \\ContosoOldPrtSrv1\OldPrtQ3,\\ContosoNewPrtSrv1\NewPrtQ3               ==>      Print Queue '\\ContosoOldPrtSrv1\OldPrtQ3' will be replaced by '\\ContosoNewPrtSrv1\NewPrtQ3'  
        \\ContosoOldPrtSrv1\OldPrtQ4,                                           ==>      Print Queue '\\ContosoOldPrtSrv1\OldPrtQ4' will be univocally removed if found  
        \\ContosoOldPrtSrv1\OldPrtQ5,DELETE                                     ==>      Print Queue '\\ContosoOldPrtSrv1\OldPrtQ5' will be univocally removed if found  
        \\ContosoNewPrtSrv1\NewGrpPrtQ,INSTALL                                  ==>      Print Queue '\\ContosoNewPrtSrv1\NewGrpPrtQ' will be univocally added if not found  


# Script configuration:  
There are 5 configurable variables (see lines 149 to 153) which need to be set by IT Administrator prior to using the present Script:  
* Variable '**DeptsOU**' contains the parent node OU in the form "OU=xyz" where all departments are residing.  
* Variable '**PrintServersOU**' contains the OU where the Print Servers involved in the migration are located. Specifying this allows for fatser LDAP searches.  
* Variable '**PrintMigGroupsOU**' contains the OU where the different printer migrations Groups are residing. Specifying this allows for fatser LDAP searches.  
          Printer Migrations Groups for each BU/Department/OU are expected to match the pattern "_[SubOUNameInDeptsOUVar]-PrinterMigration_".  
          So for instance for HR, the script expects an 'HR' OU inside '**DeptsOU**' above and the AD Group containg HR users which can migrate to be named "_HR-PrinterMigration_" and reside in AD inside the OU specified through this "**PrintMigGroupsOU**" variable.  
* Variable '**PrintQMappingsRepo**' contains the UNC location of where the mapping file(s) reside. (Ensure NTFS Security and Share permissions are set for 'Everyone' to READ).  
          Printer Mappings files to be posted here are expected to match the string pattern "_<SubOUNameIn{DeptsOU}Var>-PrintQMig.csv_".  
          So for instance again for HR, the script expects an 'HR' OU inside '**DeptsOU**' above and the file containing the mappings for HR must be called "_HR-PrintQMig.csv_" and obviously must be present on the Network Share specified in this '**PrintQMappingsRepo**'.  
* Variable '**PrintQMigrationLogs**' contains the UNC location of where the user Print Queue migrations logs are to be created. (Ensure NTFS Security and Share permissions are set for 'Everyone' to WRITE).  


Note: The bahaviour for user's Default Printer is that if for whatever reason no new Default Printer can be set in its place, it will never be removed.  
