<#

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
#>

[CmdletBinding()]
PARAM
(
	[Parameter(Mandatory=$true,HelpMessage='Please enter the FQDN of the SCCM Site Server to use')][Alias('SCCMConnection')][String]$SiteServer,
	[Parameter(Mandatory=$true,HelpMessage='Please enter the FQDN of the SCOM Management Server to use')][Alias('SCOMConnection')][String]$ScomServer,
    [Parameter(Mandatory=$true,HelpMessage='Please enter the name of the output Management Pack (i.e., "MCM.Maintenance.Windows.Management.Pack"')][Alias('ManagementPack')][String]$mp

)

Function ConvertFrom-CCMSchedule {
    <#
    .SYNOPSIS
        Convert Configuration Manager Schedule Strings
    .DESCRIPTION
        This function will take a Configuration Manager Schedule String and convert it into a readable object, including
        the calculated description of the schedule
    .PARAMETER ScheduleString
        Accepts an array of strings. This should be a schedule string in the MEMCM format
    .EXAMPLE
        PS C:\> ConvertFrom-CCMSchedule -ScheduleString 1033BC7B10100010
        SmsProviderObjectPath : SMS_ST_RecurInterval
        DayDuration : 0
        DaySpan : 2
        HourDuration : 2
        HourSpan : 0
        IsGMT : False
        MinuteDuration : 59
        MinuteSpan : 0
        StartTime : 11/19/2019 1:04:00 AM
        Description : Occurs every 2 days effective 11/19/2019 1:04:00 AM
    .NOTES
        This function was created to allow for converting MEMCM schedule strings without relying on the SDK / Site Server
        It also happens to be a TON faster than the Convert-CMSchedule cmdlet and the CIM method on the site server
    #>
    Param(
        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
        [Alias('Schedules')]
        [string[]]$ScheduleString
    )
    begin {
        #region TypeMap for returning readable window type
        $TypeMap = @{
            1 = 'SMS_ST_NonRecurring'
            2 = 'SMS_ST_RecurInterval'
            3 = 'SMS_ST_RecurWeekly'
            4 = 'SMS_ST_RecurMonthlyByWeekday'
            5 = 'SMS_ST_RecurMonthlyByDate'
            6 = 'SMS_ST_RecurMonthlyByWeekdayBase'
        }
        #endregion TypeMap for returning readable window type

        #region function to return a formatted day such as 1st, 2nd, or 3rd
        function Get-FancyDay {
            <#
                .SYNOPSIS
                Convert the input 'Day' integer to a 'fancy' value such as 1st, 2nd, 4d, 4th, etc.
            #>
            param(
                [int]$Day
            )
            $Suffix = switch -regex ($Day) {
                '1(1|2|3)$' {
                    'th'
                    break
                }
                '.?1$' {
                    'st'
                    break
                }
                '.?2$' {
                    'nd'
                    break
                }
                '.?3$' {
                    'rd'
                    break
                }
                default {
                    'th'
                    break
                }
            }
            [string]::Format('{0}{1}', $Day, $Suffix)
        }
        #endregion function to return a formatted day such as 1st, 2nd, or 3rd
    }
    process {
        # we will split the schedulestring input into 16 characters, as some are stored as multiple in one
        foreach ($Schedule in ($ScheduleString -split '(\w{16})' | Where-Object { $_ })) {
            $MW = [ordered]@{ }

            # the first 8 characters are the Start of the MW, while the last 8 characters are the recurrence schedule
            $Start = $Schedule.Substring(0, 8)
            $Recurrence = $Schedule.Substring(8, 8)
            # Convert to binary string and pad left with 0 to ensure 32 character length for consistent parsing
            $binaryRecurrence = [Convert]::ToString([int64]"0x$Recurrence".ToString(), 2).PadLeft(32, 48)

            [bool]$IsGMT = [Convert]::ToInt32($binaryRecurrence.Substring(31, 1), 2)

            switch ($Start) {
                '00012000' {
                    # this is as 'simple' schedule, such as a CI that 'runs once a day' or 'every 8 hours'
                }
                default {
                    # Convert to binary string and pad left with 0 to ensure 32 character length for consistent parsing
                    $binaryStart = [Convert]::ToString([int64]"0x$Start".ToString(), 2).PadLeft(32, 48)

                    # Collect timedata and ensure we pad left with 0 to ensure 2 character length
                    [string]$StartMinute = ([Convert]::ToInt32($binaryStart.Substring(0, 6), 2).ToString()).PadLeft(2, 48)
                    [string]$MinuteDuration = [Convert]::ToInt32($binaryStart.Substring(26, 6), 2).ToString()
                    [string]$StartHour = ([Convert]::ToInt32($binaryStart.Substring(6, 5), 2).ToString()).PadLeft(2, 48)
                    [string]$StartDay = ([Convert]::ToInt32($binaryStart.Substring(11, 5), 2).ToString()).PadLeft(2, 48)
                    [string]$StartMonth = ([Convert]::ToInt32($binaryStart.Substring(16, 4), 2).ToString()).PadLeft(2, 48)
                    [String]$StartYear = [Convert]::ToInt32($binaryStart.Substring(20, 6), 2) + 1970

                    # set our StartDateTimeObject variable by formatting all our calculated datetime components and piping to Get-Date
                    $Kind = switch ($IsGMT) {
                        $true {
                            [DateTimeKind]::Utc
                        }
                        $false {
                            [DateTimeKind]::Local
                        }
                    }
                    $StartDateTimeObject = [datetime]::new($StartYear, $StartMonth, $StartDay, $StartHour, $StartMinute, '00', $Kind)
                }
            }

            <#
                Day duration is found by calculating how many times 24 goes into our TotalHourDuration (number of times being DayDuration)
                and getting the remainder for HourDuration by using % for modulus
            #>
            $TotalHourDuration = [Convert]::ToInt32($binaryRecurrence.Substring(0, 5), 2)

            switch ($TotalHourDuration -gt 24) {
                $true {
                    $Hours = $TotalHourDuration % 24
                    $DayDuration = ($TotalHourDuration - $Hours) / 24
                    $HourDuration = $Hours
                }
                $false {
                    $HourDuration = $TotalHourDuration
                    $DayDuration = 0
                }
            }

            $RecurType = [Convert]::ToInt32($binaryRecurrence.Substring(10, 3), 2)

            $MW['SmsProviderObjectPath'] = $TypeMap[$RecurType]
            $MW['DayDuration'] = $DayDuration
            $MW['HourDuration'] = $HourDuration
            $MW['MinuteDuration'] = $MinuteDuration
            $MW['IsGMT'] = $IsGMT
            $MW['StartTime'] = $StartDateTimeObject

            Switch ($RecurType) {
                1 {
                    $MW['Description'] = [string]::Format('Occurs on {0}', $StartDateTimeObject)
                }
                2 {
                    $MinuteSpan = [Convert]::ToInt32($binaryRecurrence.Substring(13, 6), 2)
                    $Hourspan = [Convert]::ToInt32($binaryRecurrence.Substring(19, 5), 2)
                    $DaySpan = [Convert]::ToInt32($binaryRecurrence.Substring(24, 5), 2)
                    if ($MinuteSpan -ne 0) {
                        $Span = 'minutes'
                        $Interval = $MinuteSpan
                    }
                    elseif ($HourSpan -ne 0) {
                        $Span = 'hours'
                        $Interval = $HourSpan
                    }
                    elseif ($DaySpan -ne 0) {
                        $Span = 'days'
                        $Interval = $DaySpan
                    }

                    $MW['Description'] = [string]::Format('Occurs every {0} {1} effective {2}', $Interval, $Span, $StartDateTimeObject)
                    $MW['MinuteSpan'] = $MinuteSpan
                    $MW['HourSpan'] = $Hourspan
                    $MW['DaySpan'] = $DaySpan
                }
                3 {
                    $Day = [Convert]::ToInt32($binaryRecurrence.Substring(13, 3), 2)
                    $WeekRecurrence = [Convert]::ToInt32($binaryRecurrence.Substring(16, 3), 2)
                    $MW['Description'] = [string]::Format('Occurs every {0} weeks on {1} effective {2}', $WeekRecurrence, $([DayOfWeek]($Day - 1)), $StartDateTimeObject)
                    $MW['Day'] = $Day
                    $MW['ForNumberOfWeeks'] = $WeekRecurrence
                }
                4 {
                    $Day = [Convert]::ToInt32($binaryRecurrence.Substring(13, 3), 2)
                    $ForNumberOfMonths = [Convert]::ToInt32($binaryRecurrence.Substring(16, 4), 2)
                    $WeekOrder = [Convert]::ToInt32($binaryRecurrence.Substring(20, 3), 2)
                    $WeekRecurrence = switch ($WeekOrder) {
                        0 {
                            'Last'
                        }
                        default {
                            $(Get-FancyDay -Day $WeekOrder)
                        }
                    }
                    $MW['Description'] = [string]::Format('Occurs the {0} {1} of every {2} months effective {3}', $WeekRecurrence, $([DayOfWeek]($Day - 1)), $ForNumberOfMonths, $StartDateTimeObject)
                    $MW['Day'] = $Day
                    $MW['ForNumberOfMonths'] = $ForNumberOfMonths
                    $MW['WeekOrder'] = $WeekOrder
                }
                5 {
                    $MonthDay = [Convert]::ToInt32($binaryRecurrence.Substring(13, 5), 2)
                    $MonthRecurrence = switch ($MonthDay) {
                        0 {
                            # $Today = [datetime]::Today
                            # [datetime]::DaysInMonth($Today.Year, $Today.Month)
                            'the last day'
                        }
                        default {
                            "day $PSItem"
                        }
                    }
                    $ForNumberOfMonths = [Convert]::ToInt32($binaryRecurrence.Substring(18, 4), 2)
                    $MW['Description'] = [string]::Format('Occurs {0} of every {1} months effective {2}', $MonthRecurrence, $ForNumberOfMonths, $StartDateTimeObject)
                    $MW['ForNumberOfMonths'] = $ForNumberOfMonths
                    $MW['MonthDay'] = $MonthDay
                }
                6 {
                    $Day = [Convert]::ToInt32($binaryRecurrence.Substring(13, 3), 2)
                    $ForNumberOfMonths = [Convert]::ToInt32($binaryRecurrence.Substring(16, 4), 2)
                    $WeekOrder = [Convert]::ToInt32($binaryRecurrence.Substring(20, 3), 2)
                    $Offset = [Convert]::ToInt32($binaryRecurrence.Substring(24, 2), 2)

                    $WeekRecurrence = switch ($WeekOrder) {
                        0 {
                            'Last'
                        }
                        default {
                            $(Get-FancyDay -Day $WeekOrder)
                        }
                    }
                    $MW['Description'] = [string]::Format('Occurs {4} days after the {0} {1} of every {2} months effective {3}', $WeekRecurrence, $([DayOfWeek]($Day - 1)), $ForNumberOfMonths, $StartDateTimeObject,$Offset)
                    $MW['Day'] = $Day
                    $MW['ForNumberOfMonths'] = $ForNumberOfMonths
                    $MW['WeekOrder'] = $WeekOrder
                    $MW['WeekRecurrence'] = $WeekRecurrence
                    $MW['Offset'] = $Offset
                }

                Default {
                    Write-Error "Parsing Schedule String resulted in invalid type of $RecurType"
                }
            }

            [pscustomobject]$MW
        }
    }
}

Function Get-NthDayofMonth {
    [cmdletbinding()]
    param(
        # Validate the occurrence range
        [Parameter(Mandatory = $True, Position = 0)]
        [ValidateRange(0,4)]
        [Int]$NthOccurrence,
        # Validate the day of the week
        [Parameter(Mandatory = $True, Position = 1)]
        [ValidateSet("Sunday","Monday","Tuesday","Wednesday","Thursday","Friday","Saturday")]
        [String]$DayOfWeek,
        # Validate and set month (default current month)
        [ValidateRange(1,12)]
        [Int]$Month = (Get-Date).Month,
        # Validate and set year (default current year)
        [ValidateRange(1900,2100)]
        [Int]$Year = (Get-Date).Year
    )
    Begin {
        # Each occurrence can ONLY fall within specific dates.  We set those ranges here.
        Write-Verbose "Getting `$NthOccurrence ($NthOccurrence) range"
        # COULD just do $(($NthOccurrence*7)-6)..$($NthOccurrence*7) but I like seeing the ranges in the code
        $Nth = switch ($NthOccurrence) {
            1  {1..7}
            2  {8..14}
            3  {15..21}
            4  {22..28}
            0  {29..31}
        }
        Write-Verbose "`$NthOccurrence range is: $($Nth[0])..$($Nth[6])"
        Write-Verbose "Getting `$Month and `$Year date"
        # Get a DateTime object for the selected Month/Year
        $MonthYear = [DateTime]::new($Year, $Month, 1)
        Write-Verbose $MonthYear
    }
    Process {
        Write-Verbose "Getting occurrence $NthOccurrence of weekday $DayOfWeek in $($MonthYear.ToLongDateString().Split(',')[1].Trim().Split(' ')[0]) of $Year"
        # Get the day of the week for each date in range and select the desired one
        foreach ($Day in $Nth) {
            Get-Date $MonthYear -Day $Day | Where-Object {$_.DayOfWeek -eq $DayOfWeek}
        }
    }
    End {
    }
}

Function Write-ToMP {
    <#
    .SYNOPSIS
        Write out the leafs of XML to generate a fully functioning MP containing custom
        #groups with explicit members.
    .DESCRIPTION
        This function will generate a Management Pack with groups contianing explicit members.
    .PARAMETER GN
        The name of the group to create (i.e. My.SCOM.Group)
    .PARAMETER GDN
        The display name of the group to create (i.e. My SCOM Group)
    .PARAMETER M
        An array list of SCOM ObjectIds to be members of the group.
    .PARAMETER MP
        The name intended for the output Management Pack (i.e., My.Management.Pack)
    .EXAMPLE
        Write-ToMP -GN <Group.Full.Name> -GDN <Group Display Name> -M <ARRAYOFSCOMOBJECTIDS> -MP <C:\My Management Packs\My Script Folder\My.Management.Pack.xml>
      .NOTES
        This function was created specifically to create a large number of groups in a management pack based on a list of groups to create.
    #>

    Param(
        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
        [Alias('GN')]
        [string[]]$GroupName,
        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
        [Alias('GDN')]
        [string[]]$GroupDisplayName,
        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
        [Alias('M')]
        [string[]]$Members,
        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
        [Alias('MP')]
        [string[]]$managementpack
    )
    
    #Load the MP
    $xml = New-Object XML
    $xml.load($managementpack)
    #Load the namespace so we can refer to leafs explicitly
    $ns = New-Object System.Xml.XmlNamespaceManager($xml.NameTable)
    $ns.AddNamespace("ns", $xml.DocumentElement.NamespaceURI)

    ####START ADD CLASS TYPE####
    $node = $xml.SelectSingleNode("//ns:ClassTypes", $ns)
    $ClassType = $xml.CreateElement('ClassType')
    $ClassType.SetAttribute('ID',"$GroupName")
    $ClassType.SetAttribute('Accessibility',"Public")
    $ClassType.SetAttribute('Abstract',"false")
    $ClassType.SetAttribute('Base',"SystemCenter!Microsoft.SystemCenter.ComputerGroup")
    $ClassType.SetAttribute('Hosted',"false")
    $ClassType.SetAttribute('Singleton',"true")
    $ClassType.SetAttribute('Extension',"false")
    $node.AppendChild($ClassType) | Out-Null
    ####END ADD CLASS TYPE####

    ####START ADD DISCOVERY####
    #Create the discovery node for this Discovery...
    $node = $xml.SelectSingleNode("//ns:Discoveries", $ns)
    $Discovery = $xml.CreateElement('Discovery')
    $Discovery.SetAttribute('ID',"$GroupName.Discovery")
    $Discovery.SetAttribute('Enabled',"true")
    $Discovery.SetAttribute('Target',"$GroupName")
    $Discovery.SetAttribute('ConfirmDelivery',"false")
    $Discovery.SetAttribute('Remotable',"true")
    $Discovery.SetAttribute('Priority',"Normal")
    $node.AppendChild($Discovery) | Out-Null

    $parent = $xml.SelectSingleNode("//ManagementPack/Monitoring/Discoveries/Discovery[@ID='$GroupName.Discovery']")
    #Create the Category...
    $Category = $xml.CreateElement('Category')
    $Category.InnerText = "Discovery"
    $parent.AppendChild($Category) | Out-Null
#Add the Discovery Type. Since this is static, and has a sub-node, it made most sense to just import static XML, rather than add each node independently...
$DiscoveryTypes = [xml]@"
        <DiscoveryTypes>
            <DiscoveryRelationship TypeID="SystemCenter!Microsoft.SystemCenter.ComputerGroupContainsComputer" />
        </DiscoveryTypes>
"@
    $parent.AppendChild($xml.ImportNode($DiscoveryTypes.DiscoveryTypes, $true)) | Out-Null
    
    #Add the DataSource node....
    $DataSource = $xml.CreateElement('DataSource')
    $DataSource.SetAttribute('ID',"$GroupName.DiscoveryPopulator")
    $DataSource.SetAttribute('TypeID',"SystemCenter!Microsoft.SystemCenter.GroupPopulator")
    $parent.AppendChild($DataSource) | Out-Null

    #Add the RuleID inside it...
    $RuleId = $xml.CreateElement('RuleId')
    $RuleId.InnerText = "`$MPElement`$"
    $parent.DataSource.AppendChild($RuleId) | Out-Null

    #Add the GroupInstanceID...
    $GroupInstanceId = $xml.CreateElement('GroupInstanceId')
    $GroupInstanceId.InnerText = "`$MPElement[Name=`"$GroupName`"]`$"
    $parent.DataSource.AppendChild($GroupInstanceId) | Out-Null

#Add the MembershipRule. There were a bunch of ways to do this here. I thought about just passing in an array of ObjectIDs, but then I would need two loops;
#one to create the array of ObjectIDs, and one to create the same list with the additional XML elements...
foreach($member in $members){
    $xmlmembers += '<MonitoringObjectId>' + $member + '</MonitoringObjectId>'
}
$MembershipRules = [xml]@"
            <MembershipRules>
            <MembershipRule>
                <MonitoringClass>`$MPElement[Name="Windows!Microsoft.Windows.Computer"]`$</MonitoringClass>
                <RelationshipClass>`$MPElement[Name="SystemCenter!Microsoft.SystemCenter.ComputerGroupContainsComputer"]`$</RelationshipClass>
                <IncludeList>
                    $xmlmembers
                </IncludeList>
            </MembershipRule>
            </MembershipRules>
"@
    $parent.DataSource.AppendChild($xml.ImportNode($MembershipRules.MembershipRules, $true)) | Out-Null
    ####END ADD DISCOVERY####

    ####START LANGUAGE PACKS####
    $node = $xml.SelectSingleNode("//ns:LanguagePack[@ID='ENU']/DisplayStrings", $ns)
    $DisplayString1 = $xml.CreateElement('DisplayString')
    $DisplayString1.SetAttribute('ElementID',"$GroupName")
    $node.AppendChild($DisplayString1) | Out-Null

    $node = $xml.SelectSingleNode("//ns:LanguagePack[@ID='ENU']/DisplayStrings/DisplayString[@ElementID='$GroupName']", $ns)
    $Name = $xml.CreateElement('Name')
    $Name.InnerText = "$GroupDisplayName"
    $node.AppendChild($Name) | Out-Null


    $node = $xml.SelectSingleNode("//ns:LanguagePack[@ID='ENU']/DisplayStrings", $ns)
    $DisplayString2 = $xml.CreateElement('DisplayString')
    $DisplayString2.SetAttribute('ElementID',"$GroupName.Discovery")
    $node.AppendChild($DisplayString2) | Out-Null

    $node = $xml.SelectSingleNode("//ns:LanguagePack[@ID='ENU']/DisplayStrings/DisplayString[@ElementID='$GroupName.Discovery']", $ns)
    $Name = $xml.CreateElement('Name')
    $Name.InnerText = "$GroupDisplayName Discovery"
    $node.AppendChild($Name) | Out-Null

    $node = $xml.SelectSingleNode("//ns:LanguagePack[@ID='ENU']/DisplayStrings/DisplayString[@ElementID='$GroupName.Discovery']", $ns)
    $Description = $xml.CreateElement('Description')
    $Description.InnerText = "Group populator for $GroupDisplayName"
    $node.AppendChild($Description) | Out-Null
    ####END LANGUAGE PACKS####

#Save the MP
$xml.Save($managementpack)

}

#Add relevant modules, establish relevant connections
Import-Module OperationsManager
Write-Debug "Connection to Management Server $ScomServer..."
New-SCOMManagementGroupConnection -ComputerName $scomserver

#Pull the site database server and the database name for this site from the registry on the SCCM site server.
Write-Debug "Fetching database information from SCCM site server..."
$SiteCodes = Invoke-Command -ComputerName $SiteServer {Get-ChildItem -Path 'HKLM:\SOFTWARE\Microsoft\SMS\Providers\Sites'}
$SiteCode = ($SiteCodes[0].Name).Replace('HKEY_LOCAL_MACHINE\','HKLM:\') + "\"
$MCMDBServer = (Invoke-Command -ComputerName $SiteServer -ScriptBlock {Get-ItemProperty -Path $Using:SiteCode -Name 'SQL Server Name'}).'SQL Server Name'
$MCMDBName = (Invoke-Command -ComputerName $SiteServer -ScriptBlock {Get-ItemProperty -Path $Using:SiteCode -Name 'Database Name'}).'Database Name'

####COPY STUB MANAGERMENT PACK FOR EDITING####

#Copy the stub mp to a new mp for editing...
$filetocopy = (Get-Location).Path + '\BASEMP.xml'
$filetowrite =  (Get-Location).Path + '\' + $mp + '.xml'
Copy-Item $filetocopy -Destination $filetowrite -Force
$version = (Get-SCOMManagementPack  | Select-Object Name,Version | Where-Object {$_.Name -eq "$mp"}).Version

#If a management pack of the same name is already installed, increment the version number of the one we are about to create...
####BE VERY CAREFUL####
if($null -ne $version){
    $xml = New-Object XML
    $xml.load($filetowrite)
    $revision = ($version.Revision + 1)
    $newversion = [string]$version.Major +  '.' + [string]$version.Minor + '.' + [string]$version.Build  + '.' + [string]$revision
    $node = $xml.ManagementPack.Manifest.Identity
    $node.Version = $newversion
    $xml.Save($filetowrite)
}

$xml = New-Object XML
$xml.load($filetowrite)
$ns = New-Object System.Xml.XmlNamespaceManager($xml.NameTable)
$ns.AddNamespace("ns", $xml.DocumentElement.NamespaceURI)#Load the namespace so we can refer to leafs explicitly

$node = $xml.ManagementPack.Manifest.Identity
$node.ID = $mp

$node = $xml.SelectSingleNode("//ns:LanguagePack[@ID='ENU']/DisplayStrings", $ns)
$DisplayString = $xml.CreateElement('DisplayString')
$DisplayString.SetAttribute('ElementID',"$mp")
$node.AppendChild($DisplayString) | Out-Null

$node = $xml.SelectSingleNode("//ns:LanguagePack[@ID='ENU']/DisplayStrings/DisplayString[@ElementID='$mp']", $ns)
$Name = $xml.CreateElement('Name')
$DisplayName = $mp.Replace('.',' ')
$Name.InnerText = "$DisplayName"
$node.AppendChild($Name) | Out-Null

$node = $xml.SelectSingleNode("//ns:LanguagePack[@ID='ENU']/DisplayStrings/DisplayString[@ElementID='$mp']", $ns)
$Description = $xml.CreateElement('Description')
$Description.InnerText = "This Management Pack contains all auto-generated groups for Configuration Manager Collections"
$node.AppendChild($Description) | Out-Null

$xml.Save($filetowrite)


####END COPY STUB MANAGERMENT PACK FOR EDITING####
#Build a comma-separated list of all the systems in SCOM
#Pass this list in as a filter in the SQL query below, to ensure we only retrieve systems
#from MCM that are also in SCOM (and only servers)
#This also allows for multiple SCOM Management Groups served by a single MCM deployment
Write-Debug "Creating list of SCOM Systems for Maintenance Schedules..."
$scomlist = (Get-SCOMClass -Name 'Microsoft.Windows.Computer' | Get-SCOMClassInstance | Select-Object ID,DisplayName,@{Expression={$_.'[Microsoft.Windows.Computer].NetbiosComputerName'};Label="NetBiosName"}) | Where-Object {$_.DisplayName -notin (Get-SCOMManagementServer).DisplayName}
$filterlist = ''
foreach($system in $scomlist){
    $scomcomputer = $system.NetBiosName
    $filterlist = $filterlist + "'" + $scomcomputer + "'" +  ','
}
$mcmfilter = $filterlist.TrimEnd(',')

#Build a SQL query to pull systems, schedules, and collections from MCM based on the system name in SCOM
$schedulequery = @"
SELECT
v_FullCollectionMembership.Name AS Computername,
v_R_System.Full_Domain_Name0 AS Domain,
v_Collection.Name AS CollectionName,
v_Collection.CollectionId AS CollectionId,
vSMS_ServiceWindow.Name AS ScheduleName,
vSMS_ServiceWindow.Description,
vSMS_ServiceWindow.StartTime,
vSMS_ServiceWindow.Duration,
vSMS_ServiceWindow.RecurrenceType,
vSMS_ServiceWindow.ServiceWindowType,
vSMS_ServiceWindow.Schedules,
vSMS_ServiceWindow.ServiceWindowID
FROM vSMS_ServiceWindow
inner join v_FullCollectionMembership on (v_FullCollectionMembership.CollectionID = vSMS_ServiceWindow.SiteID)
inner join v_Collection on (v_Collection.CollectionID = v_FullCollectionMembership.CollectionID)
inner join v_R_System on (v_R_System.ResourceID = v_FullCollectionMembership.ResourceID)
WHERE vSMS_ServiceWindow.Enabled = 1 
    AND vSMS_ServiceWindow.RecurrenceType <> 1 
    AND v_Collection.MemberCount > 0 
    AND v_R_System.Operating_System_Name_and0 NOT LIKE '%Workstation%' 
    AND v_R_System.Full_Domain_Name0 <> ''
    AND v_FullCollectionMembership.Name IN ($mcmfilter)
ORDER BY v_Collection.Name
"@

#Execute the query...
Write-Debug "Executing SCCM SQL Database Query..."
$mcmlist = Invoke-Sqlcmd -ServerInstance $MCMDBServer -Database $MCMDBName -Query $schedulequery

#At this point, we have a list of all servers in MCM that exist in SCOM. The next step is to create a custom array
#that contains all the needed information in a single array for creating all the schedules and the new MP
Write-Debug "Building tailored array for use in creating Maintenance Schedules..."
$actionlist = New-Object System.Collections.ArrayList
foreach($object in $mcmlist){
    $fqdn = $object.Computername + '.' + $object.Domain
    $objectid = ($scomlist | Select-Object ID,DisplayName | Where-Object {$_.DisplayName -eq $fqdn}).ID.Guid
    $groupfullname = "$mp" + '.' + $object.CollectionId
    $groupdisplayname = "Configuration Manager Collection - " + $object.CollectionName
    $theschedule = ConvertFrom-CCMSchedule $object.Schedules
    $y = New-Object PSCustomObject
    $y | Add-Member -MemberType NoteProperty -Name FQDN -Value $fqdn
    $y | Add-Member -MemberType NoteProperty -Name ObjectID -Value $objectid
    $y | Add-Member -MemberType NoteProperty -Name GroupFullName -Value $groupfullname
    $y | Add-Member -MemberType NoteProperty -Name GroupDisplayName -Value $groupdisplayname
    $y | Add-Member -MemberType NoteProperty -Name ScheduleName -Value $object.ScheduleName
    $y | Add-Member -MemberType NoteProperty -Name Description -Value $object.Description
    $y | Add-Member -MemberType NoteProperty -Name StartTime -Value $object.StartTime
    $y | Add-Member -MemberType NoteProperty -Name Duration -Value $object.Duration
    $y | Add-Member -MemberType NoteProperty -Name RecurrenceType -Value $object.RecurrenceType
    $y | Add-Member -MemberType NoteProperty -Name ServiceWindowType -Value $object.ServiceWindowType
    $y | Add-Member -MemberType NoteProperty -Name Schedules -Value $theschedule
    $y | Add-Member -MemberType NoteProperty -Name ScheduleID -Value $object.ServiceWindowId
    $actionlist.Add($y) | Out-Null
}

####START WRITE NEW MANAGEMENT PACK FOR GROUPS####
$GroupList = $actionlist | Select-Object GroupFullName,GroupDisplayName -Unique
$memberlist = @()
foreach($group in $GroupList){
    $memberlist = ($actionlist | Select-Object ObjectID,GroupFullName | Where-Object {$_.GroupFullName -eq $group.GroupFullName}).ObjectID
    Write-ToMP -GN $group.GroupFullName -GDN $group.GroupDisplayName -M $memberlist -MP $filetowrite
}

#Import the generated Management Pack
Import-SCManagementPack $filetowrite

Write-Host "Sleeping for three minutes to allow processing of the Management Pack to complete..."
Start-Sleep -Seconds 180
####END WRITE NEW MANAGEMENT PACK FOR GROUPS####

####START CREATE SCHEDULES####
$schedulelist = $actionlist | Select-Object ScheduleID,Schedules,GroupFullName,GroupDisplayName,ScheduleName -Unique
Write-Debug "Begin looping through each collection and schedule..."

#We have to create a new list of groups that has the GUID of the group we want to put in recursive Maintenance Mode. The GUID does not exist until after the group is created.
$cmgrouplist = Get-SCOMClass -DisplayName "Group" | Get-SCOMClassInstance | Select-Object FullName,Id | Where-Object {$_.FullName -like "$mp.*"}

#Get a list of the existing schedules
$scomschedules = Get-SCOMMaintenanceScheduleList | Select-Object ScheduleName,ScheduleID | Where-Object {$_.ScheduleName -like "Configuration Manager Collection - *"}

#Loop through each schedule in the schedules list, and begin configuring SCOM
foreach($schedule in $schedulelist){

    #Set the schedule type. Can be one of the objects in the SWITCH below
    $schedtype = $schedule.Schedules.SmsProviderObjectPath
    
    #Set the comments
    $comments = $schedule.Schedules.Description

    #Set the ID of the group to apply the Maintenance Schedule to
    $groupid = ($cmgrouplist | Where-Object {$_.FullName -eq $schedule.GroupFullName}).Id

    #Set the duration on the Maintenance Window/Schedule
    $scomduration = (($schedule.Schedules.DayDuration * 1440) + ($schedule.Schedules.HourDuration * 60) + $schedule.Schedules.MinuteDuration)

    #Set the start time of the Maintenance Windows/Schedule (SCOM handles these slightly differently than SCCM)
    if($schedule.Schedules.IsGMT){
        $DateTime = Get-Date -Date $schedule.Schedules.StartTime
        $localtime = $DateTime.AddHours((Get-TimeZone).BaseUtcOffset.Hours)
    }else{
        $localtime = $schedule.Schedules.StartTime
    }

    #Make sure each schedule for this group has a unique name.
    $scommaintsched = $schedule.GroupDisplayName + ' ' + $schedule.ScheduleName

    SWITCH($schedtype){
        #Simple weekly schedule occurs every n number of weeks on a given day, at a given time, for a length of time
        'SMS_ST_RecurWeekly' {
            $freqtype = 8
            #We have to translate between the different way that these are handled between SCCM and SCOM
            SWITCH($schedule.Schedules.Day){
                1{$FreqInterval = 1}
                2{$FreqInterval = 2}
                3{$FreqInterval = 4}
                4{$FreqInterval = 8}
                5{$FreqInterval = 16}
                6{$FreqInterval = 32}
                7{$FreqInterval = 64}
            }

            #ForNumberOfWeeks is misleading - in SCCM this is "every 2 weeks" or "every 4 weeks" etc, not how long the schedule is to remain valid.
            $freqrecurinterval = $schedule.Schedules.ForNumberOfWeeks

            #If the Maintenance Schedule already exists, edit it instead of creating it
            #It sort of makes sense to just edit it every time we see the same schedule, as otherwise we need to add a lot of code to detemine what, if anything, has changed in the schedule
            #It's less code and less processing to just set them every time to whatever they already are
            if($scomsched = ($scomschedules | Where-Object {$_.ScheduleName -eq $scommaintsched})){
                $schedid = $scomsched.ScheduleID
                Write-Debug "SCOM Maintenance Schedule $scommaintsched already exists, so updating..."
                Edit-SCOMMaintenanceSchedule -ScheduleID $schedid -Name $scommaintsched -MonitoringObjects $groupid -ActiveStartTime $localtime -Duration $scomduration -ReasonCode PlannedOther -FreqType $freqtype -FreqInterval $FreqInterval -FreqRecurrenceFactor $freqrecurinterval -Comments $comments
            }else{
                Write-Debug "Creating SCOM Maintenance Schedule $scommaintsched..."
                New-SCOMMaintenanceSchedule -Name $scommaintsched -MonitoringObjects $groupid -ActiveStartTime $localtime -Duration $scomduration -ReasonCode PlannedOther -FreqType $freqtype -FreqInterval $FreqInterval -FreqRecurrenceFactor $freqrecurinterval -Recursive -Comments $comments | Out-Null
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
            $FreqInterval = $schedule.Schedules.MonthDay

            #ForNumberOfWeeks is misleading - in SCCM this is "every 2 weeks" or "every 4 months" etc, not how long the schedule is to remain valid.
            $FreqRecurrenceFactor = $schedule.Schedules.ForNumberOfMonths

            #If the Maintenance Schedule already exists, edit it instead of creating it
            #It sort of makes sense to just edit it every time we see the same schedule, as otherwise we need to add a lot of code to detemine what, if anything, has changed in the schedule
            #It's less code and less processing to just set them every time to whatever they already are
            if($scomsched = ($scomschedules | Where-Object {$_.ScheduleName -eq $scommaintsched})){
                $schedid = $scomsched.ScheduleID
                Write-Debug "SCOM Maintenance Schedule $scommaintsched already exists, so updating..."
                Edit-SCOMMaintenanceSchedule -ScheduleID $schedid -Name $scommaintsched -MonitoringObjects $groupid -ActiveStartTime $localtime -Duration $scomduration -ReasonCode PlannedOther -FreqType $FreqType -FreqInterval $FreqInterval -FreqRecurrenceFactor $FreqRecurrenceFactor
            }else{
                Write-Debug "Creating SCOM Maintenance Schedule $scommaintsched..."
                New-SCOMMaintenanceSchedule -Name $scommaintsched -MonitoringObjects $groupid -ActiveStartTime $localtime -Duration $scomduration -ReasonCode PlannedOther -FreqType $FreqType -FreqInterval $FreqInterval -FreqRecurrenceFactor $FreqRecurrenceFactor -Recursive -Comments $comments | Out-Null
            }
        }
        #Occurs on the nth weekday variable of the month every x number of months and lasts for the determined period of time
        'SMS_ST_RecurMonthlyByWeekday' {
                $FreqType = 32
                $FreqInterval = $schedule.Schedules.Day

                #We have to translate between the different way that these are handled between SCCM and SCOM
                SWITCH($schedule.Schedules.WeekOrder){
                    0{$FreqRelativeInterval = 16}
                    1{$FreqRelativeInterval = 1}
                    2{$FreqRelativeInterval = 2}
                    3{$FreqRelativeInterval = 4}
                    4{$FreqRelativeInterval = 8}
                }

            #ForNumberOfWeeks is misleading - in SCCM this is "every 2 weeks" or "every 4 weeks" etc, not how long the schedule is to remain valid.
            $FreqRecurrenceFactor = $schedule.Schedules.ForNumberOfMonths

            #If the Maintenance Schedule already exists, editi it instead of create it
            #It sort of makes sense to just edit it every time we see the same schedule, as otherwise we need to add a lot of code to detemine what, if anything, has changed in the schedule
            #It's less code and less processing to just set them every time to whatever they already are
            if($scomsched = ($scomschedules | Where-Object {$_.ScheduleName -eq $scommaintsched})){
                $schedid = $scomsched.ScheduleID
                Write-Debug "SCOM Maintenance Schedule $scommaintsched already exists, so updating..."
                Edit-SCOMMaintenanceSchedule -ScheduleID $schedid -Name $scommaintsched -MonitoringObjects $groupid -ActiveStartTime $localtime -Duration $scomduration -ReasonCode PlannedOther -FreqType $FreqType -FreqInterval $FreqInterval -FreqRecurrenceFactor $FreqRecurrenceFactor -FreqRelativeInterval $FreqRelativeInterval
            }else{
                Write-Debug "Creating SCOM Maintenance Schedule $scommaintsched..."
                New-SCOMMaintenanceSchedule -Name $scommaintsched -MonitoringObjects $groupid -ActiveStartTime $localtime -Duration $scomduration -ReasonCode PlannedOther -FreqType $FreqType -FreqInterval $FreqInterval -FreqRecurrenceFactor $FreqRecurrenceFactor -FreqRelativeInterval $FreqRelativeInterval -Recursive -Comments $comments | Out-Null
            }
        }
        #Occurs on the nth weekday variable of the month every x number of months and lasts for the determined period of time
        'SMS_ST_RecurMonthlyByWeekdayBase' {
            #Since this is an offset, we can only schedule it every time we run the script for the current month, so just set it to run on a specific
            #date, rather than repeating every month on a specific relative interval.
            $FreqType = 1

            #Relative day from SCCM. See https://learn.microsoft.com/en-us/mem/configmgr/develop/reference/core/servers/configure/sms_st_recurmonthlybyweekday-server-wmi-class.
            $relativeday = $schedule.Schedules.WeekOrder

            #Get the number of minutes offset from midnight that the schedule is supposed to run
            $minutes = (New-TimeSpan -Start (Get-Date $schedule.Schedules.StartTime -Format MM/dd/yyyy) -End ($schedule.Schedules.StartTime)).TotalMinutes

            #Day of the week. In SCCM for this type of schedule, 1-7 are valid, but PS is zero-based.
            #Convert to zero-based here.
            $dayofweek = [DayOfWeek] ($schedule.Schedules.Day - 1)

            #The number of days to offset the day we schedule the Maintenance Schedule.
            $offset = $schedule.Schedules.Offset

            #Get the actual date of the intended maintenance window for this month (not the relative date).
            $thedate = Get-NthDayofMonth $relativeday $dayofweek

            #Then add the number of offset days to it.
            $mmdate = $thedate.AddDays($offset)

            #The add the number of minutes offset from midnight the schedule should run.
            $mmdate = $mmdate.AddMinutes($minutes)

            #If the Maintenance Schedule already exists, editi it instead of create it
            #It sort of makes sense to just edit it every time we see the same schedule, as otherwise we need to add a lot of code to detemine what, if anything, has changed in the schedule
            #It's less code and less processing to just set them every time to whatever they already are
            if($scomsched = ($scomschedules | Where-Object {$_.ScheduleName -eq $scommaintsched})){
                $schedid = $scomsched.ScheduleID
                Write-Debug "SCOM Maintenance Schedule $scommaintsched already exists, so updating..."
                Edit-SCOMMaintenanceSchedule -ScheduleID $schedid -Name $scommaintsched -MonitoringObjects $groupid -ActiveStartTime $mmdate -Duration $scomduration -ReasonCode PlannedOther -FreqType $FreqType | Out-Null
            }else{
                Write-Debug "Creating SCOM Maintenance Schedule $scommaintsched..."
                New-SCOMMaintenanceSchedule -Name $scommaintsched -MonitoringObjects $groupid -ActiveStartTime $mmdate -Duration $scomduration -ReasonCode PlannedOther -FreqType $FreqType -Recursive -Comments $comments | Out-Null
            }
        }
        #Occurs every n days, and last for the set amoput of time
        'SMS_ST_RecurInterval' {
            $freqtype = 4
            $freqinterval = $schedule.Schedules.DaySpan
            if($scomsched = ($scomschedules | Where-Object {$_.ScheduleName -eq $scommaintsched})){
                $schedid = $scomsched.ScheduleID
                Write-Debug "SCOM Maintenance Schedule $scommaintsched already exists, so updating..."
                Edit-SCOMMaintenanceSchedule -ScheduleID $schedid -Name $scommaintsched -MonitoringObjects $groupid -ActiveStartTime $localtime -Duration $scomduration -ReasonCode PlannedOther -FreqType $freqtype -FreqInterval $freqinterval
            }else{
                Write-Debug "Creating SCOM Maintenance Schedule $scommaintsched..."
                New-SCOMMaintenanceSchedule -Name $scommaintsched -MonitoringObjects $groupid -ActiveStartTime $localtime -DurationInMinutes $scomduration -ReasonCode PlannedOther -FreqType $freqtype -FreqInterval $freqinterval -Recursive -Comments $comments | Out-Null
            }
        }
        
    }
} 
####END CREATE SCHEDULES####
