<#

.SYNOPSIS
    This script will generate System Center Operations Manager maintenance Schedules based upon Microsoft Configuration Manager Maintenance Windows.

.DESCRIPTION
    
    This script connects to your Microsoft Configuration Manager site using the ConfigurationManager PowerShell Module. It then retrieves all
    of the collections in your environment, builds an array of all those which have enabled Maintenance Windows, and then creates groups in System Center
    Operations Manager that equate to the collections (these groups will not include any Management Servers in the collections, nor any systems which
    are not being monitored by System Center Operations Manager).

    It will then add all of the computers to the groups as explicit members, then configure Maintenance Mode Schedules in System Center Operations Manager
    for each group, which are identical to the Maintenance Windows in Microsoft Configuration Manager.

    **THIS SCRIPT HAS DEPENDENCIES**

    This script requires the availability of the following PowerShell modules on the same server upon which it is executed:

    OperationsManager
    ConfigurationManager
    OpsMgrExtended - OpsMgrExtended is available freely through PowerShell Gallery: https://www.powershellgallery.com/packages/OpsMgrExtended/1.3.1

    I have tested OpsMgrExtended using SCOM 2019 UR4. The last update it had by the devleoper was for SCOM 2016. It works swimmingly in my environment,
    but your mileage may vary. I recommend robust testing using the -Debug paramter.

.PARAMETERS
    
    -SiteServer : String. Your MCM Site Server
    -ScomServer : String. Your SCOM Management Server

    This script supports the use of the -Debug parameter. This script will execute with no output to the user without use of the -Debug parameter.

.NOTES

    File Name  : Create-SCOM-Maintenance-Schedules_From-SCCM.ps1
    Author     : Scott Brown
    Appears in -full

.LINK


.EXAMPLE

    Create SCOM Maintenance Mode Schedules from SCCM Mainetenance Windows silently
    .\Create-SCOM-Maintenance-Schedules_From-SCCM.ps1 -SiteServer <MSM SITE SERVER> -ScomServer <SCOM MANAGEMENT SERVER>

.EXAMPLE

    Create SCOM Maintenance Mode Schedules from SCCM Mainetenance Windows with Debug output
    .\Create-SCOM-Maintenance-Schedules_From-SCCM.ps1 -SiteServer <MSM SITE SERVER> -ScomServer <SCOM MANAGEMENT SERVER> -Debug

.COMPONENT
    Required PowerShell Modules:

    OperationsManager
    Configurationmanager
    OpsMgrExtended - OpsMgrExtended is available freely through PowerShell Gallery: https://www.powershellgallery.com/packages/OpsMgrExtended/1.3.1

#>
[CmdletBinding()]
PARAM
(
	[Parameter(Mandatory=$true,HelpMessage='Please enter the FQDN of the SCCM Site Server to use')][Alias('SCCMConnection')][String]$SiteServer,
	[Parameter(Mandatory=$true,HelpMessage='Please enter the FQDN of the SCOM Management Server to use')][Alias('SCOMConnection')][String]$ScomServer

)

Import-Module ConfigurationManager
Import-Module OpsMgrExtended
Import-Module OperationsManager

If ($PSBoundParameters['Debug']) {
    $DebugPreference = 'Continue'
}

Write-Debug "Retrieving the Site Code from the SCCM Site Server $SiteServer..."
#Use a remote CIM query to get the Site Code of the site we're pulling from
try{
    $query = "SELECT * FROM SMS_ProviderLocation WHERE Machine LIKE '" + $SiteServer + "'"
    $siteinfo = Get-CIMInstance -Namespace “root\SMS” -Query $query -ComputerName $SiteServer -ErrorAction Stop
    $SiteCode = $siteinfo.SiteCode     
}
catch
    [System.UnauthorizedAccessException]{Write-Warning -Message “Access denied”;break}
catch
    [System.Exception] {Write-Warning -Message “$($_.Exception.Message). Line: $($_.InvocationInfo.ScriptLineNumber)” ; break}

#Once we have the Site Code, we can map the SCCM PSDrive object
Write-Debug "Creating PSDrive for $SiteCode..."
try{
    if ((Get-PSDrive -Name $SiteCode -ErrorAction SilentlyContinue | Measure-Object).Count -ne 1){
        New-PSDrive -Name $SiteCode -PSProvider CMSite -Root $SiteServer -ErrorAction Stop -Verbose:$false | Out-Null
    } 
}
catch
    [System.UnauthorizedAccessException]{Write-Warning -Message “Access denied”;break}
catch
    [System.Exception] {Write-Warning -Message “$($_.Exception.Message). Line: $($_.InvocationInfo.ScriptLineNumber)” ; break}

Write-Debug "Setting location to PSDrive for $SiteCode..."
#Set the current location to the PSDrive we just mapped
$CurrentLocation = (Get-Location).Path
Set-Location -Path $SiteCode":" -ErrorAction Stop -Verbose:$false

# Disable Fast parameter usage check for Lazy properties 
$CMPSSuppressFastNotUsedCheck = $true

Write-Debug "Fetching list of SCCM Collections..."
#Get a list of all the custom-created collections.
$collectionlist = Get-CMDeviceCollection | Where-Object {$_.IsBuiltIn -eq $false}

Write-Debug "Begin building array of collections, member computers, and Maintenance Windows..."
#Loop through the list of collections with enabled Maintenance Windows, and populate an array with a Name for each SCOM group, 
#a Display Name for each SCOM Group, the computers to be in the group, and the Maintenance Windows we will generate as Maintenance Schedules in SCOM
#Computers and Schedules will be nested array lists
$sclist = New-Object System.Collections.ArrayList
foreach($collectionid in $collectionlist){
    #get enabled Maintenance Windows for this collection
    $mw = Get-CMMaintenanceWindow -CollectionID $collectionid.CollectionID | Where-Object {$_.IsEnabled}

    #if we found an enabled schedule for this collection, add it to the list to be added to SCOM, as long as the collection contains devices
    if($null -ne $mw){
        $computers = Get-CMCollectionMember -CollectionId $collectionid.CollectionID
        if($null -ne $computers){
            $groupname = $collectionid.CollectionID + '.Group'
            $schedules = Convert-CMSchedule -ScheduleString $mw.ServiceWindowSchedules
            $groupdisplayname = $collectionid.Name
            Write-Debug "Building object for $groupdisplayname"
            $y = New-Object PSCustomObject
            $y | Add-Member -MemberType NoteProperty -Name GroupName -Value $groupname
            $y | Add-Member -MemberType NoteProperty -Name GroupDisplayName -Value $groupdisplayname
            $y | Add-Member -MemberType NoteProperty -Name Computers -Value $computers.Name
            $y | Add-Member -MemberType NoteProperty -Name Schedules -Value $schedules
            $sclist.Add($y) | Out-Null
        }

    }
}

Write-Debug "Completed array of collections."
#Close out of SCCM
Set-Location -Path $CurrentLocation

Write-Debug "Connection to Management Server $ScomServer..."
#Connect to SCOM for standard PowerShell stuff
New-SCOMManagementGroupConnection -ComputerName $ScomServer

Write-Debug "Creating unsealed Management Pack `'SCCM Maintenance Windows Management Pack`'..."
#Create the unsealed managed pack to store the groups in if it does not already exist
if(!(Get-SCOMManagementPack -Name 'SCCM.Maintenance.Windows.Management.Pack')){
    New-OMManagementPack -SDK $ScomServer -Name 'SCCM.Maintenance.Windows.Management.Pack' -DisplayName 'SCCM Maintenance Windows Management Pack' -Description "This Management Pack contains all auto-generated groups for Configuration Manager Collections" | Out-Null
}

Write-Debug "Begin looping through each collection and schedule..."
#Loop through each schedule in the schedules list, and begin configuring SCOM
foreach($object in $sclist){
    
    #Create the Computer Group in SCOM
    $scGroupName = $object.GroupName
    $scGroupDisplayName = "Configuration Manager Collection - " + $object.GroupDisplayName

    #Prepend the management pack name to the group, since New-OMComputerGroup does this automatically, we need to get the "real" name of the group to add to it later
    $fullscGroupName = 'SCCM.Maintenance.Windows.Management.Pack' + '.' + $scGroupName

    #If the Computer Group does not exist, create it, and wait until we can query for it
    if(!(Get-SCOMGroup -DisplayName $scGroupDisplayName)){
        Write-Debug "Creating SCOM Group $scGroupDisplayName..."
        New-OMComputerGroup -SDK $ScomServer -MPName 'SCCM.Maintenance.Windows.Management.Pack' -ComputerGroupName $scGroupName -ComputerGroupDisplayName $scGroupDisplayName | Out-Null
        Do {
            Write-Debug "Waiting for group $scGroupDisplayName creation to complete..."
            start-sleep -Seconds 5
        } Until (Get-SCOMGroup -DisplayName $scGroupDisplayName)

    }
    
    #Create a list out of $object.Computers
    $complist = $object | Select -ExpandProperty Computers

    #Add each computer as an explicit group member, validating:
    #1) That the computer from SCCM exists in SCOM
    #2) That the computer is NOT a Management Server (so as not to put any Management Servers in Maintenance Mode based on the collection membership)
    foreach($computer in $complist){
        $fqdns = $null
        #If the computer is not a Management Server...
        if(!(Get-SCOMManagementServer | Where-Object {$_.DisplayName -match $computer})){
            #Try to retrieve the computer from SCOM (SCCM returns single-label names anyway, so we need to get an FQDN)...
            $fqdns = Get-SCOMClass -Name 'Microsoft.Windows.Computer' | Get-SCOMClassInstance | Select-Object DisplayName,@{Expression={$_.'[Microsoft.Windows.Computer].PrincipalName'};Label="PrincipalName"} | Where-Object {$_.DisplayName -match "$computer"}
            #If the computer is found in SCOM, try to add it to the group
            if($null -ne $fqdns){
                $pn = $fqdns.PrincipalName
                Write-Debug "Adding $pn to $scGroupDisplayName (if it is not already in the group)..."
                #New-OMComputerGroupExplicitMember is fairly forgiving in that it does not bomb if the computer is already in the group
                #This lets us edit the group without doing an extraordinary amount of validation before we attempt it - this code will work for both updates and new additions
                New-OMComputerGroupExplicitMember -SDK $ScomServer -GroupName $fullscGroupName -ComputerPrincipalName $pn -WarningAction SilentlyContinue | Out-Null
            }
        }
    }

    #At this point we have all the groups created and populated, all that remains is to create the SCOM Maintenance Schedules

    #Create a list out of $object.Schedules
    $schedlist = $object | Select -ExpandProperty Schedules
    
    #Instantiate a counter to place a suffix on each schedule generated for a specific group
    $count = 1
    foreach($schedule in $schedlist){
        #Set the schedule type. Can be one of the objects in the SWITCH below
        #each schedule type has a different set of members, so we have to add them on a per-schedule type basis
        #some fields are universal, set those here
        $schedtype = $schedule.SmsProviderObjectPath

        #Set the duration on nthe Maintenance Window/Schedule
        $scomduration = (($schedule.DayDuration * 1440) + ($schedule.HourDuration * 60) + $schedule.MinuteDuration)

        #Set the start time of the Maintenance Windows/Schedule (SCOM handles these slightly differently than SCCM)
        if($schedule.IsGMT){
            $DateTime = Get-Date -Date $schedule.StartTime
            $localtime = $DateTime.AddHours((Get-TimeZone).BaseUtcOffset.Hours)
        }else{
            $localtime = $schedule.StartTime
        }

        #Make sure each schedule for this group has a unique name. We are not using the window names from SCCM.
        #SCCM Maintenance Windows are a Collection-level piece of data, and can have non-unique names across collections
        #SCOM needs to have the Maintenance Schedules associated to a group, so it just made sense to name them based on the collection group, 
        #and incremenet a numeric value. The other option would be to determine the SSCM maintenance window name, 
        #and ensure we create a unique name by concatenating the SCCM window name and the SCOM group name, which seems like overkill, and could generate
        #extraordinarily lengthy SCOM schedule names
        if($count -gt 1){
            $scommaintsched = $object.GroupDisplayName + ' ' + $count
        }else{
            $scommaintsched = $object.GroupDisplayName
        }

        
        #Set a variable for the values of the group we just created for these objects. This will be used to apply the schedules to the group.
        $collectiongroup = Get-SCOMClass -DisplayName "Group" | Get-SCOMClassInstance | Where-Object {$_.FullName -eq $fullscGroupName}

        SWITCH($schedtype){
            #Simple weekly schedule occurs every n number of weeks on a given day, at a given time, for a length of time
            'SMS_ST_RecurWeekly' {
                $freqtype = 8
                #We have to translate between the different way that these are handled between SCCM and SCOM
                SWITCH($schedule.Day){
                    1{$freqint = 1}
                    2{$freqint = 2}
                    3{$freqint = 4}
                    4{$freqint = 8}
                    5{$freqint = 16}
                    6{$freqint = 32}
                    7{$freqint = 64}
                }

            #ForNumberOfWeeks is misleading - in SCCM this is "every 2 weeks" or "every 4 weeks" etc, not how long the schedule is to remain valid.
            $freqrecurinterval = $schedule.ForNumberOfWeeks

            #If the Maintenance Schedule already exists, edit it instead of creating it
            #It sort of makes sense to just edit it every time we see the same schedule, as otherwise we need to add a lot of code to detemine what, if anything, 
            #has changed in the schedule
            #It's less code and less processing to just set them every time to whatever they already are
            if($scomsched = Get-SCOMMaintenanceScheduleList | Select-Object ScheduleName,ScheduleID | Where-Object {$_.ScheduleName -eq $scommaintsched}){
                $schedid = $scomsched.ScheduleID
                Write-Debug "SCOM Maintenance Schedule $scommaintsched already exists, so updating..."
                Edit-SCOMMaintenanceSchedule -ScheduleID $schedid -Name $scommaintsched -MonitoringObjects $collectiongroup.Id -ActiveStartTime $localtime -Duration $scomduration -ReasonCode PlannedOther -FreqType $freqtype -FreqInterval $freqint -FreqRecurrenceFactor $freqrecurinterval
            }else{
                Write-Debug "Creating SCOM Maintenance Schedule $scommaintsched..."
                New-SCOMMaintenanceSchedule -Name $scommaintsched -MonitoringObjects $collectiongroup.Id -ActiveStartTime $localtime -Duration $scomduration -ReasonCode PlannedOther -FreqType $freqtype -FreqInterval $freqint -FreqRecurrenceFactor $freqrecurinterval -Recursive | Out-Null
            }

            }
            #Non-recurring - chose to ignore one-time only schedules
            'SMS_ST_NonRecurring' {
                #This being non-Recurring, I am going to ignore it, since the intent is to capture schedules that repeat.
                $freqtype = 1
                Write-Debug "The schedule is SMS_ST_NonRecurring: $schedtype"
                
            }
            #Simple monthly schedule, occurs on the n day of month, every x months, and lasts for the length of time
            'SMS_ST_RecurMonthlyByDate' {
                $FreqType = 16
                #We have to translate between the different way that these are handled between SCCM and SCOM
                $FreqInterval = $schedule.MonthDay

                #ForNumberOfWeeks is misleading - in SCCM this is "every 2 weeks" or "every 4 months" etc, not how long the schedule is to remain valid.
                $FreqRecurrenceFactor = $schedule.ForNumberOfMonths

                #If the Maintenance Schedule already exists, edit it instead of creating it
                #It sort of makes sense to just edit it every time we see the same schedule, as otherwise we need to add a lot of code to detemine what, if anything, has changed in the schedule
                #It's less code and less processing to just set them every time to whatever they already are
                if($scomsched = Get-SCOMMaintenanceScheduleList | Select-Object ScheduleName,ScheduleID | Where-Object {$_.ScheduleName -eq $scommaintsched}){
                    $schedid = $scomsched.ScheduleID
                    Write-Debug "SCOM Maintenance Schedule $scommaintsched already exists, so updating..."
                    Edit-SCOMMaintenanceSchedule -ScheduleID $schedid -Name $scommaintsched -MonitoringObjects $collectiongroup.Id -ActiveStartTime $localtime -Duration $scomduration -ReasonCode PlannedOther -FreqType $FreqType -FreqInterval $FreqInterval -FreqRecurrenceFactor $FreqRecurrenceFactor
                }else{
                    Write-Debug "Creating SCOM Maintenance Schedule $scommaintsched..."
                    New-SCOMMaintenanceSchedule -Name $scommaintsched -MonitoringObjects $collectiongroup.Id -ActiveStartTime $localtime -Duration $scomduration -ReasonCode PlannedOther -FreqType $FreqType -FreqInterval $FreqInterval -FreqRecurrenceFactor $FreqRecurrenceFactor -Recursive | Out-Null
                }
            }
            #Occurs on the nth weekday variable of the month every x number of months and lasts for the determined period of time
            'SMS_ST_RecurMonthlyByWeekday' {
                    $FreqType = 32
                    $FreqInterval = $schedule.Day

                    #We have to translate between the different way that these are handled between SCCM and SCOM
                    SWITCH($schedule.WeekOrder){
                        0{$FreqRelativeInterval = 16}
                        1{$FreqRelativeInterval = 1}
                        2{$FreqRelativeInterval = 2}
                        3{$FreqRelativeInterval = 4}
                        4{$FreqRelativeInterval = 8}
                    }

                #ForNumberOfWeeks is misleading - in SCCM this is "every 2 weeks" or "every 4 weeks" etc, not how long the schedule is to remain valid.
                $FreqRecurrenceFactor = $schedule.ForNumberOfMonths

                #If the Maintenance Schedule already exists, editi it instead of create it
                #It sort of makes sense to just edit it every time we see the same schedule, as otherwise we need to add a lot of code to detemine what, if anything, has changed in the schedule
                #It's less code and less processing to just set them every time to whatever they already are
                if($scomsched = Get-SCOMMaintenanceScheduleList | Select-Object ScheduleName,ScheduleID | Where-Object {$_.ScheduleName -eq $scommaintsched}){
                    $schedid = $scomsched.ScheduleID
                    Write-Debug "SCOM Maintenance Schedule $scommaintsched already exists, so updating..."
                    Edit-SCOMMaintenanceSchedule -ScheduleID $schedid -Name $scommaintsched -MonitoringObjects $collectiongroup.Id -ActiveStartTime $localtime -Duration $scomduration -ReasonCode PlannedOther -FreqType $FreqType -FreqInterval $FreqInterval -FreqRecurrenceFactor $FreqRecurrenceFactor -FreqRelativeInterval $FreqRelativeInterval
                }else{
                    Write-Debug "Creating SCOM Maintenance Schedule $scommaintsched..."
                    New-SCOMMaintenanceSchedule -Name $scommaintsched -MonitoringObjects $collectiongroup.Id -ActiveStartTime $localtime -Duration $scomduration -ReasonCode PlannedOther -FreqType $FreqType -FreqInterval $FreqInterval -FreqRecurrenceFactor $FreqRecurrenceFactor -FreqRelativeInterval $FreqRelativeInterval -Recursive | Out-Null
                }
            }
            #Occurs every n days, and last for the set amoput of time
            'SMS_ST_RecurInterval' {
                $freqtype = 4
                if($scomsched = Get-SCOMMaintenanceScheduleList | Select-Object ScheduleName,ScheduleID | Where-Object {$_.ScheduleName -eq $scommaintsched}){
                    $schedid = $scomsched.ScheduleID
                    Write-Debug "SCOM Maintenance Schedule $scommaintsched already exists, so updating..."
                    Edit-SCOMMaintenanceSchedule -ScheduleID $schedid -Name $scommaintsched -MonitoringObjects $collectiongroup.Id -ActiveStartTime $localtime -Duration $scomduration -ReasonCode PlannedOther -FreqType $freqtype -FreqInterval $schedule.DaySpan
                }else{
                    Write-Debug "Creating SCOM Maintenance Schedule $scommaintsched..."
                    New-SCOMMaintenanceSchedule -Name $scommaintsched -MonitoringObjects $collectiongroup.Id -ActiveStartTime $localtime -DurationInMinutes $scomduration -ReasonCode PlannedOther -FreqType $freqtype -FreqInterval $schedule.DaySpan -Recursive | Out-Null
                }
            }
        
        
        }
    $count++
    }

}

Write-Debug "Script Completed."