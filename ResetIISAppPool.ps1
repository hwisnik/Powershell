<#
.SYNOPSIS
    Reset IIS Application Pool
.DESCRIPTION
    This script resets specified IIS application pool on multiple remote servers
    Designed for interactive and automation.  Interactive users can call script 
    without specifying an app pool and will be provided a choice of app pools
.PARAMETER ApplicationPoolName
    The name of the application pool as shown in InetMgr Application Pool Screen
    If you don't provide an application pool name, a Powershell GridView will 
    allow you to choose an app pool
.PARAMETER StopOnError
    Intended for use with automation, valid app pool is expected to be supplied
    as a parameter.  If app pool validation fails, the script exits 
    
.EXAMPLE
    Interactive user that doesn't know app pool name. User will be presented a 
    list of app pools to choose from:
    PS M:\>c:<CRLF>
    PS C:\>cd dev\Powershell<CRLF>
    PS C:\dev\Powershell>.\TestServersRemoteResetAppPoolv2.ps1<CRLF> 
.EXAMPLE
    Interactive user that supplies application pool name.
    Name will be validated and if it doesn't exist user will be presented a list
    of app pools
    PS M:\>c:<CRLF>
    PS C:\>cd dev\Powershell<CRLF>
    PS C:\dev\Powershell>.\TestServersRemoteResetAppPoolv2.ps1 -applicationPoolName "GTEDSAccountData-test.gtfc.com"<CRLF> 
.EXAMPLE
    Automation user (ex: Jenkins)
    C:\PS> c:\dev\Powershell <CRLF>
    C:\dev\Powershell> c:\dev\Powershell\TestServersRemoteResetAppPoolv2.ps1 -applicationPoolName "GTEDSAccountData-test.gtfc.com" -StopOnError <CRLF>     
    
.NOTES
    Author: Howard Wisnik
    Date:   29 December 2017
    Script requires companion file named HostFile.txt that identifies the list of remote servers to 
    run on.  HostFile.txt must be saved in the same directory as the ResetIISAppPool.ps1 script  .  
    ResetIISAppPool will output a logfile: AppPoolResetResults.txt with results and/or any errors in the
    same directory as the ResetIISAppPool.ps1 is located.
#>



#Declare input parameters first
param(
[Parameter(Mandatory = $false, HelpMessage="Call this script with -ApplicationPoolName poolName.  Ex: ./ResetIISAppPool.ps1 -ApplicationPoolName  `"GTEDSAccountData-test.gtfc.com`"")]
[string]$applicationPoolName,
[Parameter(Mandatory = $false, HelpMessage = "StopOnError is intended to be used with automation. Stops execution if AppPool not found.  Call this script with 2 switches -ApplicationPoolName and -StopOnError.  Ex: ./scriptname.ps1 -ApplicationPoolName  `"GTEDSAccountData-test.gtfc.com`" - StopOnError")]
[switch]$StopOnError
)

Function GetRemoteAppPools
{
[cmdletbinding()]
param (
    [ValidateNotNullorEmpty()]
    [Parameter(Mandatory = $true, HelpMessage="Internal script error. Default remote server name missing")]
    [string]$remoteServer)

    $null = @(
        #old syntax
        #$processes = get-WmiObject Win32_process -ComputerName $REMOTE_SERVER | where CommandLine -Match "w3wp.exe" |ForEach-Object -MemberName CommandLine 
        #new syntax - get w3wp processes on remote server.  The CommandLine property contains the app pool name
        $processes1 = Get-CimInstance -Query "SELECT CommandLine from Win32_Process WHERE CommandLine LIKE '%w3wp.exe%'" -ComputerName $remoteServer |select CommandLine -ExpandProperty CommandLine

        $appPoolArrayList = New-Object System.Collections.ArrayList

        ForEach ($pool in $processes1) 
        {
            #Sample $pool string:
            #c:\windows\system32\inetsrv\w3wp.exe -ap "GTEDSPayoffQuote-testv1_0.gtfc.com" -v "v4.0" -l "webengine4.dll" -a \\.\pipe\iisipm3b593e61-b651-4fbe-9041-e6412cad255a -h "C:\inetpub\temp\apppools\GTEDSPayoffQuote-testv1_0.gtfc.com\GTEDSPayoffQuote-testv1_0.gtfc.com.config" -w "" -m 0
            #extract from first '"' to second '"' to get the app pool

            $appPool = $pool.Substring($pool.IndexOf('"')+1, $pool.Substring($pool.indexof('"') +1).IndexOf('"')) 
            $appPoolArrayList.Add($appPool)
        }

    )
    return $appPoolArrayList
}

Function ConvertToInt64  #Returns 0 if conversion fails
{
    param (
    [string]$strToConvert)
    [int64]$convertedInt = 0
    [bool]$result = [int64]::TryParse($strToConvert,[ref]$convertedInt
    )
    if (!$result) 
        {$convertedInt = 0}
    return $convertedInt
}

Function ResetRemoteAppPool
{
    param ([string]$remoteServer, 
    [string]$applicationPoolName
    )
    
# Pipe entire function to null so that only desired variables are returned
    $null = @(
        # Pipe to Out-String is there to ensure that PS waits until recycle completes
        Invoke-Command -ComputerName "$remoteServer" -ScriptBlock { Import-Module WebAdministration;Restart-WebAppPool -Name $args[0]} -ArgumentList $applicationPoolName |Out-String

        $WsBytes = Get-WmiObject Win32_process -ComputerName $remoteServer | where CommandLine -Match $applicationPoolName |ForEach-Object -MemberName WS  | Out-String
 
        #Get-WmiObject Win32_process returns 2 WS values afer recycle First is the Working set before recycle, second is WS after recycle
        #Need to split response object (Newline delimited) into array and get First and Second value

        $WsBytesArray = $WsBytes.Split([Environment]::NewLine)

        [string]$memValue0String = $WsBytesArray |Select-Object -First 1 |Out-String
        [string]$memValue1String =  $WsBytesArray |Select-Object -Skip 1 |Out-String


        #Now convert to Int64 in a safe way
        [int64]$memValue0Int = 0
        [int64]$memValue1Int = 0

        #Convert to Int
        $memValue0Int= ConvertToInt64 $memValue0String
        $memValue1Int= ConvertToInt64 $memValue1String

        #Convert bytes to megaBytes:  2 to the 20th bytes = 1 megaByte (MB)
        $beforeMBytes = [math]::Round($memValue0Int /[math]::Pow(2,20)) 
        $afterMBytes = [math]::Round($memValue1Int /[math]::Pow(2,20)) 

        $resetDateTime = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    )
    return $beforeMBytes, $afterMBytes , $resetDateTime
}

Function ValidateAppPool
{
[cmdletbinding()]
    
Param(
    [string]$remoteServer,
    [Parameter(Mandatory = $true)]
    #[ValidateScript({GetRemoteAppPools -appPool $_ -remoteServer "tp8vucsmspt01.gtfc.com" |Out-GridView  -PassThru -Title "Choose Application Pool!" } )]
    [string]$applicationPoolName
    )
    #Check one server and assume all other servers have same app pool configuration
    $appPoolList = GetRemoteAppPools "tp8vucsmspt01.gtfc.com" 
    if (!$appPoolList.Contains($applicationPoolName))
    {
        if($StopOnError.IsPresent)  #Stop if validation fails and script invoked by Automation (Jenkins) by using -StopOnError switch on commandline
        {   Write-Output("Stopping due to AppPool not Found")
            exit
        }
        else
        {
           $applicationPoolName = $appPoolList |Sort-Object | Out-GridView  -PassThru -Title "Choose Application Pool!"
        }
    }
    return $applicationPoolName
}

Function GetFirstServerInHostsFile
{
[cmdletbinding()]
param (
    [ValidateNotNullorEmpty()]
    [Parameter(Mandatory = $true, HelpMessage="Internal script error. HostFile.txt path missing")]
    [string]$filePathAndName)


    if (!(Test-Path $filePathAndName)) 
    {
        throw [System.IO.FileNotFoundException] "$filePathAndName not found."
    }

    foreach($server in [System.IO.File]::ReadLines($filePathAndName))
    {
       if (!$server.StartsWith("#"))
       {
            $rServer = $server
            break
       }

    }

    If ([string]::IsNullOrEmpty($rServer))
    {
    # no servers found throw exception
    throw [System.Exception] "No Server names found in $filePathAndName"
    }
    return $rServer
}

#Main Application

#Assign a value of space to applicationPoolName if it is empty so that user can pick from a list
if([string]::IsNullOrEmpty($applicationPoolName)) {$applicationPoolName = " "}


try
{
#get first server from Hosts file to validate that requested app pool exists.
#assumes same app pool naming convention for all servers in HostFile.txt

$currentDir = Get-Location
$infilePathAndName = Join-Path $currentDir "\HostFile.txt"
$outfilePathAndName = Join-Path $currentDir "\AppPoolResetResults.txt"

$remoteServer = GetFirstServerInHostsFile  $infilePathAndName

#Validate ApplicationPoolName
$applicationPoolName = ValidateAppPool -remoteServer $remoteServer -applicationPoolName $applicationPoolName

Clear-Host

Write-Output("Resetting AppPool $ApplicationPoolName") 

$userName = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
$outputArrayList = New-Object System.Collections.ArrayList


foreach($remoteServer in [System.IO.File]::ReadLines($infilePathAndName))
{
       if (!$remoteServer.StartsWith("#"))
       {
        #I wanted to write $outMessage to both the console and to a file
        #Adding $outMessage to an ArrayList worked great for the file, but
        #PS was fighting me as it was outputting to the console an enumeration for each iteration through the loop
        #Used the null redirect trick to remove the enumeration
        #Then, Write-Output didn't work so I used System.Console::Write
        #However, doing all of this removed the NewLine after each iteration so I added the join of a Newline
        $null = @(
            $returnBeforeMB, $returnAfterMB, $resetDateTime = ResetRemoteAppPool $remoteServer $applicationPoolName
            [string]$outMessage = "Server: $remoteServer AppPool: $applicationPoolName WorkingSet BEFORE Recycle: $returnBeforeMB MB WorkingSet AFTER recycle: $returnAfterMB MB at $resetDateTime by $userName" #|Out-String
            [System.Console]::Write(-join ($outMessage, [Environment]::NewLine))    #used this instead of Write-Output to get rid of unwanted enumeration being di
            $OutputArrayList.Add($outMessage)
            )
       }
}

$outputArrayList | Out-File -filepath $outfilePathAndName -Width 200 

Write-Output ("Script complete")
}
Catch
{
    $_.Exception|format-list -force |Tee-Object -FilePath $outfilePathAndName 
    "ResetIISAppPool failed and stopped execution" |Tee-Object -FilePath $outfilePathAndName -Append
    Exit
}






