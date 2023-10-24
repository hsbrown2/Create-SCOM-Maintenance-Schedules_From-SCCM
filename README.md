.SYNOPSIS
    
    This script will generate System Center Operations Manager Maintenance Schedules based upon Microsoft Configuration Manager Maintenance Windows.

.DESCRIPTION
    
    This script connects to your Microsoft Configuration Manager database. 
    It then retrieves all of the collections in your environment, builds an array of all those which have enabled Maintenance Windows, and then creates groups in System Center
    Operations Manager that equate to the collections (these groups will not include any Management Servers in the collections, nor any systems which
    are not being monitored by System Center Operations Manager).

    It will then add all of the computers to the groups as explicit members, then configure Maintenance Mode Schedules in System Center Operations Manager
    for each group, which are identical to the Maintenance Windows in Microsoft Configuration Manager.

    **THIS SCRIPT HAS DEPENDENCIES**

    This script requires the availability of the following PowerShell modules on the same server upon which it is executed:

    OperationsManager

    It is also required that the stub Management Pack BASEMP.xml is contained in the folder the script executes from.
    This base MP will be used to generate the MP containing the required groups.
 

.PARAMETERS
    
    -SCCMConnection : String. Your MCM Site Server
    -SCOMConnection : String. Your SCOM Management Server
    -ManagementPack : String. The name of the output Management Pack (i.e., "MCM.Maintenance.Windows.Management.Pack"')

    This script supports the use of the -Debug parameter. This script will execute with no output to the user without use of the -Debug parameter.

.NOTES

    File Name  : Create-SCOM-Maintenance-Schedules_From-SCCM.ps1
    Author     : Scott Brown
    Appears in -full

.LINK
    
    https://github.com/hsbrown2/Create-SCOM-Maintenance-Schedules_From-SCCM

.EXAMPLE

    Create SCOM Maintenance Mode Schedules from SCCM Mainetenance Windows silently
    .\Create-SCOM-Maintenance-Schedules_From-SCCM.ps1 -SCCMConnection <MCM SITE SERVER> -SCOMConnection <SCOM MANAGEMENT SERVER> -ManagementPack <MANAGEMENT.PACK.NAME>

.EXAMPLE

    Create SCOM Maintenance Mode Schedules from SCCM Mainetenance Windows with Debug output
    .\Create-SCOM-Maintenance-Schedules_From-SCCM.ps1 -SCCMConnection <MCM SITE SERVER> -SCOMConnection <SCOM MANAGEMENT SERVER> -ManagementPack <MANAGEMENT.PACK.NAME> -Debug

.COMPONENT
    
    Required PowerShell Modules:
        OperationsManager

    Required files:
        
        BASEMP.xml