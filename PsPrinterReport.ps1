#requires -version 3
<#
.SYNOPSIS
  PsPrinterReport
.DESCRIPTION
  Creates a human readable report of all Printers on a specified print server
.PARAMETER <Parameter_Name>
    <Brief description of parameter input required. Repeat this attribute if required>
.INPUTS
  Parameters above
.OUTPUTS
  Reports stored in $sReportPath
  Logs stored in $sLogPath
.NOTES
  Version:        v0.1-Aplha
.EXAMPLE
  Just Run it
#>

#---------------------------------------------------------[Initialisations]--------------------------------------------------------

Clear-Host

#Set Error Action to Stop, if an error occurs the script wil stop!
$ErrorActionPreference = "Stop"

#Dot Source or Import required Functions and Libraries
Import-module PsLogHandler

#----------------------------------------------------------[Declarations]----------------------------------------------------------

########### Script info
$sScriptName = "Report-Printers"
$sScriptVersion = "v0.1-Aplha"
$sScriptMaintainer = "s.lootens@mzh.nl"

########### SMTP server info
$sSMTPServer = "smtp.mz.local"
$sSMTPSender = "no-reply@mz.local"

########### Log File Info
$sLogPath = "$env:USERPROFILE\Desktop\$sScriptName"
$sLogName = ((get-date).tostring("yyyy-MM-dd") + " $sScriptName.log")
$sLogFile = Join-Path -Path $sLogPath -ChildPath $sLogName

########### Report File Info
$sReportPath = "$env:USERPROFILE\Desktop\$sScriptName"

########### Time, Date vars
$Today = (get-date)

########### Other
$PrintServer = 'PR06'

#-----------------------------------------------------------[Functions]------------------------------------------------------------

#-----------------------------------------------------------[Execution]------------------------------------------------------------

#Log dir check, if not existend create the direcotry if posible.
if (!(Test-Path $sLogPath)) { 
    New-Item -path $sLogPath -ItemType Directory | Out-Null
}

#Report dir check, if not existend create the direcotry if posible.
if (!(Test-Path $sReportPath)) { 
    New-Item -path $sReportPath -ItemType Directory | Out-Null
}

#Log Start
Log-Start -LogPath $sLogPath -LogName $sLogName -ScriptVersion $sScriptVersion -ScriptName $sScriptName -ScriptMaintainer $sScriptMaintainer


ipconfig.exe /flushdns

#***************************************************************************************************
#                      Receiving Printers on Printserver.
#***************************************************************************************************

Log-Write -LineValue ('Recieving all printers from print server ' + $PrintServer) -LogPath $sLogFile
$Printers = Get-Printer -ComputerName pr06 -Full

$CustomPrinterObjects = @()

#***************************************************************************************************
#                      Procesing the recieved printers.
#***************************************************************************************************

#Filter and format the recieved printers.
Log-Write -LineValue 'Procesing the recieved printers' -LogPath $sLogFile -Header $true
foreach  ($Printer in $Printers)
{
    Log-Write -LineValue ('Procesing printer ' + $Printer.name) -LogPath $sLogFile -Header $true
    Write-Host ('Procesing printer ' + $Printer.name)

    if ($Printer.DeviceType -ne 'Print')
    {
        
        Log-Write -LineValue 'not of type Print, skipping' -LogPath $sLogFile
        continue
    }

    Log-Write -LineValue 'preprocessing some printer information' -LogPath $sLogFile
    
    $CustomPrinterObject = New-Object System.Object
    $CustomPrinterObject | % {
        $_ | Add-Member -type NoteProperty -name Name -value $Printer.Name
        $_ | Add-Member -type NoteProperty -name Location -value $Printer.Location
        $_ | Add-Member -type NoteProperty -name Comment -value $Printer.Comment
        $_ | Add-Member -type NoteProperty -name PrinterStatus -value $Printer.PrinterStatus    
        $_ | Add-Member -type NoteProperty -name DriverName -value $Printer.DriverName
        $_ | Add-Member -type NoteProperty -name Shared -value $Printer.Shared
        $_ | Add-Member -type NoteProperty -name Published -value $Printer.Published
        $_ | Add-Member -type NoteProperty -name PortName -value $Printer.PortName        
        $_ | Add-Member -type NoteProperty -name ConnectionTest -value $null
        $_ | Add-Member -type NoteProperty -name ConnectionTestDesc -value $null
        $_ | Add-Member -type NoteProperty -name IPAddress -value $null

    }

    Log-Write -LineValue ($CustomPrinterObject | Out-String) -LogPath $sLogFile
                   
    if ($Printer.PortName -match "\b\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}\b" )
    {     
        
        Log-Write -LineValue 'Printer Portname is filled with IP Adress' -LogPath $sLogFile
        $CustomPrinterObject.IPAddress = $Printer.PortName

        if (Test-Connection -ComputerName $Printer.PortName -Count 2 -TimeToLive 15 -ErrorAction SilentlyContinue)
        {

            Log-Write -LineValue 'Connection Test Succeeded' -LogPath $sLogFile
            $CustomPrinterObject.ConnectionTest = 'Succeeded'
            $CustomPrinterObject.ConnectionTestDesc = 'IP Ping Succeeded'

        }

        else
        {

            Log-Write -LineValue 'Connection Test Failed' -LogPath $sLogFile
            $CustomPrinterObject.ConnectionTest = 'Failed'
            $CustomPrinterObject.ConnectionTestDesc = 'IP Ping Failed'

        }

    }

    Else
    {
        
        Log-Write -LineValue 'Printer Portname is filled with DNS Name' -LogPath $sLogFile
        
        if (Resolve-DnsName $Printer.PortName -ErrorAction SilentlyContinue)
        {
            Log-Write -LineValue 'DNS Name Resolved' -LogPath $sLogFile
            $CustomPrinterObject.IPAddress = (Resolve-DnsName $Printer.PortName | Where-Object {$_.Type -eq "A"}).IPAddress

            if (Test-Connection -ComputerName $Printer.PortName -Count 2 -TimeToLive 15 -ErrorAction SilentlyContinue)
            {
                Log-Write -LineValue 'Connection Test Succeeded' -LogPath $sLogFile
                $CustomPrinterObject.ConnectionTest = 'Succeeded'
                $CustomPrinterObject.ConnectionTestDesc = 'DNS Ping Succeeded'

            }

            else
            {

                Log-Write -LineValue 'Connection Test Failed' -LogPath $sLogFile
                $CustomPrinterObject.ConnectionTest = 'Failed'
                $CustomPrinterObject.ConnectionTestDesc = 'DNS Ping Failed'
            
            }
        }

        Else
        {

            Log-Write -LineValue 'Could not resolve DNS Name' -LogPath $sLogFile
            Log-Write -LineValue 'Connection Test Failed' -LogPath $sLogFile

            $CustomPrinterObject.ConnectionTest = 'Failed'
            $CustomPrinterObject.ConnectionTestDesc = 'DNS Resolve Failed'
        }

    }
    $CustomPrinterObjects += $CustomPrinterObject
}


$CustomPrinterObjects | Export-Csv -Path ($sReportPath + '\'+ ((get-date).tostring("yyyy-MM-dd")) + " Report-Printers.csv") -Delimiter ';' -NoTypeInformation














#***************************************************************************************************
#                      Receiving PrintJob Related events on Printserver.
#***************************************************************************************************

#Log-Write -LineValue ('Recieving all PrintJob events from print server ' + $PrintServer) -LogPath $sLogFile
#$PrintJobEvents = Get-WinEvent Microsoft-Windows-PrintService/Operational -ComputerName $PrintServer

#***************************************************************************************************
#                      Procesing the recieved Events.
#***************************************************************************************************