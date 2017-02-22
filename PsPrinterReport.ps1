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
Import-module .\Modules\PsLogHandler.psm1

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
    $CustomPrinterObject | Add-Member -type NoteProperty -name Name -value $Printer.Name
    $CustomPrinterObject | Add-Member -type NoteProperty -name Location -value $Printer.Location
    $CustomPrinterObject | Add-Member -type NoteProperty -name Comment -value $Printer.Comment
    $CustomPrinterObject | Add-Member -type NoteProperty -name PrinterStatus -value $Printer.PrinterStatus    
    $CustomPrinterObject | Add-Member -type NoteProperty -name DriverName -value $Printer.DriverName
    $CustomPrinterObject | Add-Member -type NoteProperty -name Shared -value $Printer.Shared
    $CustomPrinterObject | Add-Member -type NoteProperty -name Published -value $Printer.Published
    $CustomPrinterObject | Add-Member -type NoteProperty -name PortName -value $Printer.PortName    



    Log-Write -LineValue ($CustomPrinterObject | Out-String) -LogPath $sLogFile
                   
    if ($Printer.PortName -match "\b\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}\b" )
    {     
        
        Log-Write -LineValue 'Printer Portname is filled with IP Adress' -LogPath $sLogFile
        
        if (Test-Connection -ComputerName $Printer.PortName -Count 2 -TimeToLive 15 -ErrorAction SilentlyContinue)
        {

            Log-Write -LineValue 'Connection Test Succeeded' -LogPath $sLogFile
            $CustomPrinterObject | Add-Member -type NoteProperty -name ConnectionTest -value Succeeded
            $CustomPrinterObject | Add-Member -type NoteProperty -name ConnectionTestDesc -value 'IP Ping Succeeded'

        }

        else
        {

            Log-Write -LineValue 'Connection Test Failed' -LogPath $sLogFile
            $CustomPrinterObject | Add-Member -type NoteProperty -name ConnectionTest -value Failed
            $CustomPrinterObject | Add-Member -type NoteProperty -name ConnectionTestDesc -value 'IP Ping Failed'

        }

    }

    Else
    {
        
        Log-Write -LineValue 'Printer Portname is filled with DNS Name' -LogPath $sLogFile

        if (Resolve-DnsName $Printer.PortName -ErrorAction SilentlyContinue)
        {
            
            Log-Write -LineValue 'DNS Name Resolved' -LogPath $sLogFile

            if (Test-Connection -ComputerName $Printer.PortName -Count 2 -TimeToLive 15 -ErrorAction SilentlyContinue)
            {

                Log-Write -LineValue 'Connection Test Succeeded' -LogPath $sLogFile
                $CustomPrinterObject | Add-Member -type NoteProperty -name ConnectionTest -value Succeeded
                $CustomPrinterObject | Add-Member -type NoteProperty -name ConnectionTestDesc -value 'DNS Ping Succeeded'

            }

            else
            {

                Log-Write -LineValue 'Connection Test Failed' -LogPath $sLogFile
                $CustomPrinterObject | Add-Member -type NoteProperty -name ConnectionTest -value Failed
                $CustomPrinterObject | Add-Member -type NoteProperty -name ConnectionTestDesc -value 'DNS Ping Failed'
            
            }
        }

        Else
        {

            Log-Write -LineValue 'Could not resolve DNS Name' -LogPath $sLogFile
            Log-Write -LineValue 'Connection Test Failed' -LogPath $sLogFile

            $CustomPrinterObject | Add-Member -type NoteProperty -name ConnectionTest -value Failed
            $CustomPrinterObject | Add-Member -type NoteProperty -name ConnectionTestDesc -value 'DNS Resolve Failed'
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