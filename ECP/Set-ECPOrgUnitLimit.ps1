Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn

# Location of the ecp directory under ClientAccess directory
$ECPConfigDir = 'C:\Program Files\Microsoft\Exchange Server\V15\ClientAccess\ecp'

# List of Exchange Servers - This may be pulled in any way desired, or specifically listed in a powershell variable. I use the DAG Command.
$Serverlist = (Get-DatabaseAvailabilityGroup "<DAGNAME>").Servers.Name

# Shall we install the features required for remote Recycling of the App Pools? Only needs to be ran once. Set to $true for installation.
$installfeatures = $false

# Thanks and Credits: 
# This script was compiled by blacktirion on GitHub. Thanks to the users below for some starting points on some of this.
# Reddit User: u/676f626c7565 for the XML Update/Add scriptblock which I slightly modified (https://www.reddit.com/r/exchangeserver/comments/8vjyju/comment/e1s532s/)
# Reddit User: u/zarberg for the Start-RecycleAppPool Function, which I copied in full. (https://www.reddit.com/r/exchangeserver/comments/8vjyju/comment/e1s7tw9/)

Function Start-RecycleAppPool
    {
    [CmdletBinding()]
    param(
        [parameter(Mandatory=$true, Position=0,ValueFromPipeLine= $true)]
        [string]$Server,
        [parameter(Mandatory=$true,Position=1)]
        [string]$apppool
        )
    process{
    if ((Test-Connection $Server -count 1 -ErrorAction SilentlyContinue) -and ($apppoolarray=gwmi -namespace "root\webadministration" -Class applicationpool -ComputerName $Server -Authentication 6 -Property name -ea 'SilentlyContinue' | select -ExpandProperty name))
    {
    if ($apppool -eq "ALL")
    {gwmi -namespace "root\webadministration" -Class applicationpool -ComputerName $Server -Authentication 6 | Invoke-WmiMethod -Name recycle -ErrorAction SilentlyContinue
    if ($?)
    {Write-Host "All apppools recycled on $server"}
    else {Write-Host -BackgroundColor Red "One apppool is either stopped or did not start backup"}
    }
    elseif ($apppool)
    {if ($apppoolarray -contains $apppool)
    {
    gwmi -namespace "root\webadministration" -ComputerName $server -Authentication 6 -Query "select * from applicationpool where name='$apppool'" | Invoke-WmiMethod -Name recycle -ErrorAction SilentlyContinue
    if ($?)
    {Write-Host "$apppool recycled on $server"}
    else
    {Write-Host "$apppool Apppool state Error"}
    }
    else
    {Write-Host -BackgroundColor Red "$apppool does not exist on $server"}}
    }
    else
    {Write-Host -BackgroundColor Red "$server not reachable or WMI Windows feature for IIS not installed"}
}
}

Function Update-WebConfig
{
    [CmdletBinding()]
    Param (
        $ECPConfigDir
    )
    $webConfig = $ECPConfigDir+'web.config'
    $webConfigbackup = $ECPConfigDir+'web.config.old'

    [xml]$doc = Get-Content $webConfig

    $obj = $doc.configuration.appSettings.add | where {$_.Key -eq 'GetListDefaultResultSize'}
    if ($obj) {
        Write-Host -BackgroundColor DarkRed "ECP Web Config Entry Already Exists for $($env:COMPUTERNAME) with value $($obj.value)... Skipping..." 
    } else {
        Write-Host -BackgroundColor DarkGreen "Updating ECP web.config for $($env:COMPUTERNAME)"
        Copy-Item $webConfig $webConfigbackup
        $newAppSetting = $doc.CreateElement("add")
        $comment = $doc.CreateComment('allows the OU picker when placing a new mailbox in its designated organizational unit to retrieve all OUs - default value is 500')
        $doc.DocumentElement.appSettings.AppendChild($comment)
        $doc.configuration.appSettings.AppendChild($newAppSetting)
        $newAppSetting.SetAttribute("key","GetListDefaultResultSize")
        $newAppSetting.SetAttribute("value","5000")
        $doc.Save($webConfig)
    }
}

if ($installfeatures -eq $true) {
    foreach ($server in $Serverlist) {
        if (((Get-WindowsFeature -ComputerName $server | where {$_.name -eq "Web-Scripting-Tools"}).InstallState -ne "Installed")){Add-WindowsFeature "Web-Scripting-Tools" -ComputerName $server}
    }
}

foreach ($server in $Serverlist) {
    $serversession = New-PSSession -ComputerName $server
    Invoke-Command -Session $serversession -ScriptBlock ${Function:Update-WebConfig} -ArgumentList $ECPConfigDir
    Remove-PSSession $serversession
    Start-RecycleAppPool -apppool MSExchangeECPAppPool -Server $server
}

