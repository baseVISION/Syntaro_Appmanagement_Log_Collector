<#
.DESCRIPTION
Collects Syntaro Appmanagement Logs

.EXAMPLE


.NOTES
Author: Pascal Berger/baseVISION
Date:   28.03.2018

History
    001: First Version


ExitCodes:
    99001: Could not Write to LogFile
    99002: Could not Write to Windows Log
    99003: Could not Set ExitMessageRegistry

#>


## Manual Variable Definition
########################################################
$VBS = new-object -comobject wscript.shell

$DebugPreference = "Continue"
$ScriptVersion = "1.0"

$Script:ThisScriptParentPath = $MyInvocation.MyCommand.Path -replace $myInvocation.MyCommand.Name,""

#If the Script gets executed as EXE we need another way to get ThisScriptParentPath
If(-not($script:ThisScriptParentPath)){
    $Script:ThisScriptParentPath = [System.Diagnostics.Process]::GetCurrentProcess() | Select-Object -ExpandProperty Path | Split-Path
}

$Script:ScriptName = $myInvocation.MyCommand.Name

$LogFilePathFolder = $env:temp
$LogFilePath = "$LogFilePathFolder\Syntaro_Appmanagement_Log_Collector.log"

# Log Configuration
$DefaultLogOutputMode  = "LogFile" # "Console-LogFile","Console-WindowsEvent","LogFile-WindowsEvent","Console","LogFile","WindowsEvent","All"
$DefaultLogWindowsEventSource = $ScriptName
$DefaultLogWindowsEventLog = "CustomPS"

$Temp = $env:temp

$CollectionFolderName = $env:COMPUTERNAME +"_" +(get-date -Format yyyyMMdd_HHmmss)

$CollectionFolder = "$Temp\$CollectionFolderName"

$SyntaroCacheFolder = "$env:ProgramData\Syntaro\Cache"
$ChacheContentFile = "$CollectionFolder\cache_content.txt"

$SyntaroLogsFolder = "$env:windir\Logs\_Syntaro"
$SyntaroLogsExportFolder = "$CollectionFolder\Logs"

$RegKey = "HKLM\SOFTWARE\Syntaro\ApplicationManagement"

$ZipFileName = "_CollectedLogs.zip"
$ZipFile ="$CollectionFolder\$ZipFileName"


#region Functions
########################################################

function Check-LogFileSize {
    <#
    .DESCRIPTION
    Check if the Logfile exceds a defined Size and if yes rolles id over to a .old.log.

    .PARAMETER Log
    Specifies the the Path to the Log.

    .PARAMETER MaxSize
    MaxSize in MB for the Maximum Log Size

    .EXAMPLE
    Check-LogFileSize -Log "C:\Temp\Super.log" -Size 1

    #>
    param(
        [Parameter(Mandatory=$true,Position=1)]
        [String]
        $Log
    ,
        [Parameter(Mandatory=$true)]
        [String]
        $MaxSize

    )

    
    #Create the old.log File
    $LogOld = $Log.Insert(($Log.LastIndexOf(".")),".old")
        
      if (Test-Path $Log) {
             Write-Log "The Log $Log exists"
        $FileSizeInMB= ((Get-ItemProperty -Path $Log).Length)/1MB
        Write-Log "The Logs Size is $FileSizeInMB MB"
        #Compare the File Size
        If($FileSizeInMB -ge $MaxSize){
            Write-Log "The definde Maximum Size is $MaxSize MB I need to rollover the Log"
            #If the old.log File already exists remove it
            if (Test-Path $LogOld) {
                Write-Log "The Rollover File $LogOld already exists. I will remove it first"
                Remove-Item -path $LogOld -Force
            }
            #Rename the Log
            Rename-Item -Path $Log -NewName $LogOld -Force
            Write-Log "Rolled the Log file over to $LogOld"

        }
        else{
            Write-Log "The definde Maximum Size is $MaxSize MB no need to rollover"
        }

      } else {
             Write-Log "The Log $Log dosen't exists"
      }
}

function Write-Log {
    <#
    .DESCRIPTION
    Write text to a logfile with the current time.

    .PARAMETER Message
    Specifies the message to log.

    .PARAMETER Type
    Type of Message ("Info","Debug","Warn","Error").

    .PARAMETER OutputMode
    Specifies where the log should be written. Possible values are "Console","LogFile" and "Both".

    .PARAMETER Exception
    You can write an exception object to the log file if there was an exception.

    .EXAMPLE
    Write-Log -Message "Start process XY"

    .NOTES
    This function should be used to log information to console or log file.
    #>
    param(
        [Parameter(Mandatory=$true,Position=1)]
        [String]
        $Message
    ,
        [Parameter(Mandatory=$false)]
        [ValidateSet("Info","Debug","Warn","Error")]
        [String]
        $Type = "Debug"
    ,
        [Parameter(Mandatory=$false)]
        [ValidateSet("Console-LogFile","Console-WindowsEvent","LogFile-WindowsEvent","Console","LogFile","WindowsEvent","All")]
        [String]
        $OutputMode = $DefaultLogOutputMode
    ,
        [Parameter(Mandatory=$false)]
        [Exception]
        $Exception
    )
    
    $DateTimeString = Get-Date -Format "yyyy-MM-dd HH:mm:sszz"
    $Output = ($DateTimeString + "`t" + $Type.ToUpper() + "`t" + $Message)
    if($Exception){
        $ExceptionString =  ("[" + $Exception.GetType().FullName + "] " + $Exception.Message)
        $Output = "$Output - $ExceptionString"
    }

    if ($OutputMode -eq "Console" -OR $OutputMode -eq "Console-LogFile" -OR $OutputMode -eq "Console-WindowsEvent" -OR $OutputMode -eq "All") {
        if($Type -eq "Error"){
            Write-Error $output
        } elseif($Type -eq "Warn"){
            Write-Warning $output
        } elseif($Type -eq "Debug"){
            Write-Debug $output
        } else{
            Write-Verbose $output -Verbose
        }
    }
    
    if ($OutputMode -eq "LogFile" -OR $OutputMode -eq "Console-LogFile" -OR $OutputMode -eq "LogFile-WindowsEvent" -OR $OutputMode -eq "All") {
        try {
            Add-Content $LogFilePath -Value $Output -ErrorAction Stop
        } catch {
            exit 99001
        }
    }

    if ($OutputMode -eq "Console-WindowsEvent" -OR $OutputMode -eq "WindowsEvent" -OR $OutputMode -eq "LogFile-WindowsEvent" -OR $OutputMode -eq "All") {
        try {
            New-EventLog -LogName $DefaultLogWindowsEventLog -Source $DefaultLogWindowsEventSource -ErrorAction SilentlyContinue
            switch ($Type) {
                "Warn" {
                    $EventType = "Warning"
                    break
                }
                "Error" {
                    $EventType = "Error"
                    break
                }
                default {
                    $EventType = "Information"
                }
            }
            Write-EventLog -LogName $DefaultLogWindowsEventLog -Source $DefaultLogWindowsEventSource -EntryType $EventType -EventId 1 -Message $Output -ErrorAction Stop
        } catch {
            exit 99002
        }
    }
}

function New-Folder{
    <#
    .DESCRIPTION
    Creates a Folder if it's not existing.

    .PARAMETER Path
    Specifies the path of the new folder.

    .EXAMPLE
    CreateFolder "c:\temp"

    .NOTES
    This function creates a folder if doesn't exist.
    #>
    param(
        [Parameter(Mandatory=$True,Position=1)]
        [string]$Path
    )
	# Check if the folder Exists

	if (Test-Path $Path) {
		Write-Log "Folder: $Path Already Exists"
	} else {
		New-Item -Path $Path -type directory | Out-Null
		Write-Log "Creating $Path"
	}
}

function Set-RegValue {
    <#
    .DESCRIPTION
    Set registry value and create parent key if it is not existing.

    .PARAMETER Path
    Registry Path

    .PARAMETER Name
    Name of the Value

    .PARAMETER Value
    Value to set

    .PARAMETER Type
    Type = Binary, DWord, ExpandString, MultiString, String or QWord

    #>
    param(
        [Parameter(Mandatory=$True)]
        [string]$Path,
        [Parameter(Mandatory=$True)]
        [string]$Name,
        [Parameter(Mandatory=$True)]
        [AllowEmptyString()]
        [string]$Value,
        [Parameter(Mandatory=$True)]
        [string]$Type
    )
    
    try {
        $ErrorActionPreference = 'Stop' # convert all errors to terminating errors
        Start-Transaction

	   if (Test-Path $Path -erroraction silentlycontinue) {      
 
        } else {
            New-Item -Path $Path -Force
            Write-Log "Registry key $Path created"  
        } 
        $null = New-ItemProperty -Path $Path -Name $Name -PropertyType $Type -Value $Value -Force
        Write-Log "Registry Value $Path, $Name, $Type, $Value set"
        Complete-Transaction
    } catch {
        Undo-Transaction
        Write-Log "Registry value not set $Path, $Name, $Value, $Type" -Type Error -Exception $_.Exception
    }
}

function Set-ExitMessageRegistry () {
    <#
    .DESCRIPTION
    Write Time and ExitMessage into Registry. This is used by various reporting scripts and applications like ConfigMgr or the OSI Documentation Script.

    .PARAMETER Scriptname
    The Name of the running Script

    .PARAMETER LogfileLocation
    The Path of the Logfile

    .PARAMETER ExitMessage
    The ExitMessage for the current Script. If no Error set it to Success

    #>
    param(
    [Parameter(Mandatory=$True)]
    [string]$Script = "$ScriptName`_$ScriptVersion`.ps1",
    [Parameter(Mandatory=$False)]
    [string]$LogfileLocation=$LogFilePath,
    [Parameter(Mandatory=$True)]
    [string]$ExitMessage
    )

    $DateTime = Get-Date –f o
    #The registry Key into which the information gets written must be checked and if not existing created
    if((Test-Path "HKLM:\SOFTWARE\_Custom") -eq $False)
    {
        $null = New-Item -Path "HKLM:\SOFTWARE\_Custom"
    }
    if((Test-Path "HKLM:\SOFTWARE\_Custom\Scripts") -eq $False)
    {
        $null = New-Item -Path "HKLM:\SOFTWARE\_Custom\Scripts"
    }
    try { 
        #The new key gets created and the values written into it
        $null = New-Item -Path "HKLM:\SOFTWARE\_Custom\Scripts\$Script" -ErrorAction Stop -Force
        $null = New-ItemProperty -Path "HKLM:\SOFTWARE\_Custom\Scripts\$Script" -Name "Scriptname" -Value "$Script" -ErrorAction Stop -Force
        $null = New-ItemProperty -Path "HKLM:\SOFTWARE\_Custom\Scripts\$Script" -Name "Time" -Value "$DateTime" -ErrorAction Stop -Force
        $null = New-ItemProperty -Path "HKLM:\SOFTWARE\_Custom\Scripts\$Script" -Name "ExitMessage" -Value "$ExitMessage" -ErrorAction Stop -Force
        $null = New-ItemProperty -Path "HKLM:\SOFTWARE\_Custom\Scripts\$Script" -Name "LogfileLocation" -Value "$LogfileLocation" -ErrorAction Stop -Force
    } catch { 
        Write-Log "Set-ExitMessageRegistry failed" -Type Error -Exception $_.Exception
        #If the registry keys can not be written the Error Message is returned and the indication which line (therefore which Entry) had the error
        exit 99003
    }
}
#endregion

#region Dynamic Variables and Parameters
########################################################


#endregion

#region Initialization
########################################################

New-Folder $LogFilePathFolder

If(Test-Path $LogFilePath){
    remove-item $LogFilePath -Force -ErrorAction SilentlyContinue
}

Write-Log "---------------------------------------------------Start Script $Scriptname---------------------------------------------------"


#endregion

#region Main Script
########################################################

New-Folder $CollectionFolder

Write-Log "Exporting Registry"
IF(Test-Path $RegKey.Replace("HKLM\","HKLM:\")){

  $process = Start-Process reg -argumentlist ("export ""HKLM\SOFTWARE\Syntaro\ApplicationManagement"" ""$CollectionFolder\Registry.txt""") -PassThru -Wait | Out-Null

  If($process.ExitCode -eq 0){
    Write-Log "Exported regristrykey"
  }
  else{
    Write-Log  ("Failed to export regristrykey / Exit Code of reg.exe was " +$process.ExitCode) -Type Error
  }

}
else{
    Write-Log  "Regkey $RegKey dosen't exist" -Type Error
    Write-Log  "The Syntaro base Package is not installed on this computer" -Type Error
    $Answer = $VBS.popup(("The Syntaro base Package is not installed on this computer!"),0,"Syntaro is not installed",16)
    Exit


}

Write-Log "Exporting Eventlogs"
$log = Get-WmiObject -Class Win32_NTEventlogFile | Where-Object LogfileName -EQ 'Syntaro'

If($log.NumberOfRecords -gt 0){
    Write-Log  "Syntaro Eventlog found"

    $EventLogBackupFile = "$CollectionFolder\AppmanagementEventlog.evtx"
    $log.BackupEventlog($EventLogBackupFile) |Out-Null

    If(Test-Path $EventLogBackupFile){
        Write-Log "Exported Eventlog"
    }
    else{
        Write-Log "Failed to export the eventlog" -Type Error
    }
}
else{
    Write-Log  "No Syntaro Eventlog found" -Type Error
}


Write-Log "Collecting Cache Folder Content"
IF(Test-Path $SyntaroCacheFolder){

    Write-Log  "Found cache folder $SyntaroCacheFolder"

    $CacheContent = Get-ChildItem -Recurse $SyntaroCacheFolder 
    $CacheContent | Out-File -FilePath $ChacheContentFile

    If(Test-Path $ChacheContentFile){
        Write-Log "Exported chache content"
    }
    else{
        Write-Log "Failed export chache content" -Type Error
    }
}
else{
    Write-Log  "No Cache folder found at $SyntaroCacheFolder" -Type Error
}


Write-Log  "Export Syntaro Logs"
IF(Test-Path $SyntaroLogsFolder){

    Write-Log  "Found Logs folder $SyntaroLogsFolder"

    $SyntaroLogsContent = Get-ChildItem -Recurse $SyntaroLogsFolder

    If($SyntaroLogsContent.count -gt 0){

        New-Folder $SyntaroLogsExportFolder

        try{
            Copy-Item -Recurse -Path $SyntaroLogsFolder -Destination $SyntaroLogsExportFolder
            Write-Log  "copied the logs to  $SyntaroCacheFolder"
        }
        catch{
            Write-Log  "failed to copy the logs to  $SyntaroCacheFolder" -Type Error -Exception $_.Exception
        }
    }
    else{
        Write-Log "Found no Logfiles in the folder $SyntaroLogsFolder" -Type Warn
    }
}
else{
    Write-Log  "No Cache folder found at $SyntaroCacheFolder" -Type Error
}

Copy-Item $LogFilePath -Destination $CollectionFolder


Compress-Archive -CompressionLevel Optimal -Path $CollectionFolder -DestinationPath $ZipFile

$process = Start-Process explorer.exe -ArgumentList $CollectionFolder -PassThru

Start-Sleep -Seconds 2
$Answer = $VBS.popup(("The Syntaro Appmanagement Logs are collected in the folder:`n$CollectionFolder`n`nI opend the folder in explorer for you.`n`nPlease send the file:`n$ZipFileName`nto:`nsupport@basevision.ch"),0,"Syntaro Appmanagement Logs Collected",64)


#endregion

#region Finishing
########################################################

Write-Log "---------------------------------------------------End Script $Scriptname---------------------------------------------------"

#endregion
# SIG # Begin signature block
# MIIXxQYJKoZIhvcNAQcCoIIXtjCCF7ICAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUQ8UdvEBqA5p8t467bAUbNSB3
# lUugghL4MIID7jCCA1egAwIBAgIQfpPr+3zGTlnqS5p31Ab8OzANBgkqhkiG9w0B
# AQUFADCBizELMAkGA1UEBhMCWkExFTATBgNVBAgTDFdlc3Rlcm4gQ2FwZTEUMBIG
# A1UEBxMLRHVyYmFudmlsbGUxDzANBgNVBAoTBlRoYXd0ZTEdMBsGA1UECxMUVGhh
# d3RlIENlcnRpZmljYXRpb24xHzAdBgNVBAMTFlRoYXd0ZSBUaW1lc3RhbXBpbmcg
# Q0EwHhcNMTIxMjIxMDAwMDAwWhcNMjAxMjMwMjM1OTU5WjBeMQswCQYDVQQGEwJV
# UzEdMBsGA1UEChMUU3ltYW50ZWMgQ29ycG9yYXRpb24xMDAuBgNVBAMTJ1N5bWFu
# dGVjIFRpbWUgU3RhbXBpbmcgU2VydmljZXMgQ0EgLSBHMjCCASIwDQYJKoZIhvcN
# AQEBBQADggEPADCCAQoCggEBALGss0lUS5ccEgrYJXmRIlcqb9y4JsRDc2vCvy5Q
# WvsUwnaOQwElQ7Sh4kX06Ld7w3TMIte0lAAC903tv7S3RCRrzV9FO9FEzkMScxeC
# i2m0K8uZHqxyGyZNcR+xMd37UWECU6aq9UksBXhFpS+JzueZ5/6M4lc/PcaS3Er4
# ezPkeQr78HWIQZz/xQNRmarXbJ+TaYdlKYOFwmAUxMjJOxTawIHwHw103pIiq8r3
# +3R8J+b3Sht/p8OeLa6K6qbmqicWfWH3mHERvOJQoUvlXfrlDqcsn6plINPYlujI
# fKVOSET/GeJEB5IL12iEgF1qeGRFzWBGflTBE3zFefHJwXECAwEAAaOB+jCB9zAd
# BgNVHQ4EFgQUX5r1blzMzHSa1N197z/b7EyALt0wMgYIKwYBBQUHAQEEJjAkMCIG
# CCsGAQUFBzABhhZodHRwOi8vb2NzcC50aGF3dGUuY29tMBIGA1UdEwEB/wQIMAYB
# Af8CAQAwPwYDVR0fBDgwNjA0oDKgMIYuaHR0cDovL2NybC50aGF3dGUuY29tL1Ro
# YXd0ZVRpbWVzdGFtcGluZ0NBLmNybDATBgNVHSUEDDAKBggrBgEFBQcDCDAOBgNV
# HQ8BAf8EBAMCAQYwKAYDVR0RBCEwH6QdMBsxGTAXBgNVBAMTEFRpbWVTdGFtcC0y
# MDQ4LTEwDQYJKoZIhvcNAQEFBQADgYEAAwmbj3nvf1kwqu9otfrjCR27T4IGXTdf
# plKfFo3qHJIJRG71betYfDDo+WmNI3MLEm9Hqa45EfgqsZuwGsOO61mWAK3ODE2y
# 0DGmCFwqevzieh1XTKhlGOl5QGIllm7HxzdqgyEIjkHq3dlXPx13SYcqFgZepjhq
# IhKjURmDfrYwggSjMIIDi6ADAgECAhAOz/Q4yP6/NW4E2GqYGxpQMA0GCSqGSIb3
# DQEBBQUAMF4xCzAJBgNVBAYTAlVTMR0wGwYDVQQKExRTeW1hbnRlYyBDb3Jwb3Jh
# dGlvbjEwMC4GA1UEAxMnU3ltYW50ZWMgVGltZSBTdGFtcGluZyBTZXJ2aWNlcyBD
# QSAtIEcyMB4XDTEyMTAxODAwMDAwMFoXDTIwMTIyOTIzNTk1OVowYjELMAkGA1UE
# BhMCVVMxHTAbBgNVBAoTFFN5bWFudGVjIENvcnBvcmF0aW9uMTQwMgYDVQQDEytT
# eW1hbnRlYyBUaW1lIFN0YW1waW5nIFNlcnZpY2VzIFNpZ25lciAtIEc0MIIBIjAN
# BgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEAomMLOUS4uyOnREm7Dv+h8GEKU5Ow
# mNutLA9KxW7/hjxTVQ8VzgQ/K/2plpbZvmF5C1vJTIZ25eBDSyKV7sIrQ8Gf2Gi0
# jkBP7oU4uRHFI/JkWPAVMm9OV6GuiKQC1yoezUvh3WPVF4kyW7BemVqonShQDhfu
# ltthO0VRHc8SVguSR/yrrvZmPUescHLnkudfzRC5xINklBm9JYDh6NIipdC6Anqh
# d5NbZcPuF3S8QYYq3AhMjJKMkS2ed0QfaNaodHfbDlsyi1aLM73ZY8hJnTrFxeoz
# C9Lxoxv0i77Zs1eLO94Ep3oisiSuLsdwxb5OgyYI+wu9qU+ZCOEQKHKqzQIDAQAB
# o4IBVzCCAVMwDAYDVR0TAQH/BAIwADAWBgNVHSUBAf8EDDAKBggrBgEFBQcDCDAO
# BgNVHQ8BAf8EBAMCB4AwcwYIKwYBBQUHAQEEZzBlMCoGCCsGAQUFBzABhh5odHRw
# Oi8vdHMtb2NzcC53cy5zeW1hbnRlYy5jb20wNwYIKwYBBQUHMAKGK2h0dHA6Ly90
# cy1haWEud3Muc3ltYW50ZWMuY29tL3Rzcy1jYS1nMi5jZXIwPAYDVR0fBDUwMzAx
# oC+gLYYraHR0cDovL3RzLWNybC53cy5zeW1hbnRlYy5jb20vdHNzLWNhLWcyLmNy
# bDAoBgNVHREEITAfpB0wGzEZMBcGA1UEAxMQVGltZVN0YW1wLTIwNDgtMjAdBgNV
# HQ4EFgQURsZpow5KFB7VTNpSYxc/Xja8DeYwHwYDVR0jBBgwFoAUX5r1blzMzHSa
# 1N197z/b7EyALt0wDQYJKoZIhvcNAQEFBQADggEBAHg7tJEqAEzwj2IwN3ijhCcH
# bxiy3iXcoNSUA6qGTiWfmkADHN3O43nLIWgG2rYytG2/9CwmYzPkSWRtDebDZw73
# BaQ1bHyJFsbpst+y6d0gxnEPzZV03LZc3r03H0N45ni1zSgEIKOq8UvEiCmRDoDR
# EfzdXHZuT14ORUZBbg2w6jiasTraCXEQ/Bx5tIB7rGn0/Zy2DBYr8X9bCT2bW+IW
# yhOBbQAuOA2oKY8s4bL0WqkBrxWcLC9JG9siu8P+eJRRw4axgohd8D20UaF5Mysu
# e7ncIAkTcetqGVvP6KUwVyyJST+5z3/Jvz4iaGNTmr1pdKzFHTx/kuDDvBzYBHUw
# ggUnMIIED6ADAgECAhAB7uHu9sSCKgqVgJoW5mKYMA0GCSqGSIb3DQEBCwUAMHIx
# CzAJBgNVBAYTAlVTMRUwEwYDVQQKEwxEaWdpQ2VydCBJbmMxGTAXBgNVBAsTEHd3
# dy5kaWdpY2VydC5jb20xMTAvBgNVBAMTKERpZ2lDZXJ0IFNIQTIgQXNzdXJlZCBJ
# RCBDb2RlIFNpZ25pbmcgQ0EwHhcNMTkwMzE0MDAwMDAwWhcNMjIwNjExMTIwMDAw
# WjBkMQswCQYDVQQGEwJDSDESMBAGA1UECBMJU29sb3RodXJuMREwDwYDVQQHDAhE
# w6RuaWtlbjEWMBQGA1UEChMNYmFzZVZJU0lPTiBBRzEWMBQGA1UEAxMNYmFzZVZJ
# U0lPTiBBRzCCASIwDQYJKoZIhvcNAQEBBQADggEPADCCAQoCggEBAMfM1Zl2D4mV
# Ej7w9rAboqVD3E6JHf3GUOw0cPPP94occ4dCeqITcVME6s2nhVcnff+68FPtJB2g
# BKWIB8zL4bD1SZBgLywRe3F/KvmbULw9gp5Qk8nLeVOLtXsyKEIfNMzMWeMxTMsx
# mtr910G0knpBnuHQgJVNpKF4BgSpIJZ8FQJlvYvLm0y73HXj/YSUJt7bstqnJ9Q6
# s+ngp/en1pykXhzgj76u6yPKc/kdZQwfzsLj2FQ3y7ScWt7Ps1fevkh8JBmJc+ti
# 6oKVDMArOEj7IdXn9rjPkeTSakFoqb1ceNRnMYyMOEaflMFwxylT2NTm4cYTF65m
# HtZEm//K7QECAwEAAaOCAcUwggHBMB8GA1UdIwQYMBaAFFrEuXsqCqOl6nEDwGD5
# LfZldQ5YMB0GA1UdDgQWBBTP/wNYXuPPmMSxVMf7zj/HD4TQbTAOBgNVHQ8BAf8E
# BAMCB4AwEwYDVR0lBAwwCgYIKwYBBQUHAwMwdwYDVR0fBHAwbjA1oDOgMYYvaHR0
# cDovL2NybDMuZGlnaWNlcnQuY29tL3NoYTItYXNzdXJlZC1jcy1nMS5jcmwwNaAz
# oDGGL2h0dHA6Ly9jcmw0LmRpZ2ljZXJ0LmNvbS9zaGEyLWFzc3VyZWQtY3MtZzEu
# Y3JsMEwGA1UdIARFMEMwNwYJYIZIAYb9bAMBMCowKAYIKwYBBQUHAgEWHGh0dHBz
# Oi8vd3d3LmRpZ2ljZXJ0LmNvbS9DUFMwCAYGZ4EMAQQBMIGEBggrBgEFBQcBAQR4
# MHYwJAYIKwYBBQUHMAGGGGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBOBggrBgEF
# BQcwAoZCaHR0cDovL2NhY2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0U0hBMkFz
# c3VyZWRJRENvZGVTaWduaW5nQ0EuY3J0MAwGA1UdEwEB/wQCMAAwDQYJKoZIhvcN
# AQELBQADggEBAJV9TChoYXmc7/3Qc8p9EgcF4I+6cnLSBpQOfYGi3f3bBcATTPle
# cJiruue4AxIpNwtGwZnqneGxAWT4C98yjBQbCq7nt0k3HF1LjeTNdExNx6cVGF4S
# 9HvcSsoqFNnQuMpzjnFredIP0LPvLQouYtEKcvYJDmx0Bb/72anpAlUiY0WzBF6t
# cYU//dNDqiw/0uQqFMuKzUTKSgGlf+bLsebz8XNIcJPrrui3dduig20oOR2V60Yq
# PmZhLDS3CXNvKXZbo/ib02zendVAFYDIoKZOtmNOalBoWwlQRYWmwQWuSJRbQFl1
# ridyuVRrTlCgWJj4547jqr4/cxmFLV+hrZcwggUwMIIEGKADAgECAhAECRgbX9W7
# ZnVTQ7VvlVAIMA0GCSqGSIb3DQEBCwUAMGUxCzAJBgNVBAYTAlVTMRUwEwYDVQQK
# EwxEaWdpQ2VydCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xJDAiBgNV
# BAMTG0RpZ2lDZXJ0IEFzc3VyZWQgSUQgUm9vdCBDQTAeFw0xMzEwMjIxMjAwMDBa
# Fw0yODEwMjIxMjAwMDBaMHIxCzAJBgNVBAYTAlVTMRUwEwYDVQQKEwxEaWdpQ2Vy
# dCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xMTAvBgNVBAMTKERpZ2lD
# ZXJ0IFNIQTIgQXNzdXJlZCBJRCBDb2RlIFNpZ25pbmcgQ0EwggEiMA0GCSqGSIb3
# DQEBAQUAA4IBDwAwggEKAoIBAQD407Mcfw4Rr2d3B9MLMUkZz9D7RZmxOttE9X/l
# qJ3bMtdx6nadBS63j/qSQ8Cl+YnUNxnXtqrwnIal2CWsDnkoOn7p0WfTxvspJ8fT
# eyOU5JEjlpB3gvmhhCNmElQzUHSxKCa7JGnCwlLyFGeKiUXULaGj6YgsIJWuHEqH
# CN8M9eJNYBi+qsSyrnAxZjNxPqxwoqvOf+l8y5Kh5TsxHM/q8grkV7tKtel05iv+
# bMt+dDk2DZDv5LVOpKnqagqrhPOsZ061xPeM0SAlI+sIZD5SlsHyDxL0xY4PwaLo
# LFH3c7y9hbFig3NBggfkOItqcyDQD2RzPJ6fpjOp/RnfJZPRAgMBAAGjggHNMIIB
# yTASBgNVHRMBAf8ECDAGAQH/AgEAMA4GA1UdDwEB/wQEAwIBhjATBgNVHSUEDDAK
# BggrBgEFBQcDAzB5BggrBgEFBQcBAQRtMGswJAYIKwYBBQUHMAGGGGh0dHA6Ly9v
# Y3NwLmRpZ2ljZXJ0LmNvbTBDBggrBgEFBQcwAoY3aHR0cDovL2NhY2VydHMuZGln
# aWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElEUm9vdENBLmNydDCBgQYDVR0fBHow
# eDA6oDigNoY0aHR0cDovL2NybDQuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJl
# ZElEUm9vdENBLmNybDA6oDigNoY0aHR0cDovL2NybDMuZGlnaWNlcnQuY29tL0Rp
# Z2lDZXJ0QXNzdXJlZElEUm9vdENBLmNybDBPBgNVHSAESDBGMDgGCmCGSAGG/WwA
# AgQwKjAoBggrBgEFBQcCARYcaHR0cHM6Ly93d3cuZGlnaWNlcnQuY29tL0NQUzAK
# BghghkgBhv1sAzAdBgNVHQ4EFgQUWsS5eyoKo6XqcQPAYPkt9mV1DlgwHwYDVR0j
# BBgwFoAUReuir/SSy4IxLVGLp6chnfNtyA8wDQYJKoZIhvcNAQELBQADggEBAD7s
# DVoks/Mi0RXILHwlKXaoHV0cLToaxO8wYdd+C2D9wz0PxK+L/e8q3yBVN7Dh9tGS
# dQ9RtG6ljlriXiSBThCk7j9xjmMOE0ut119EefM2FAaK95xGTlz/kLEbBw6RFfu6
# r7VRwo0kriTGxycqoSkoGjpxKAI8LpGjwCUR4pwUR6F6aGivm6dcIFzZcbEMj7uo
# +MUSaJ/PQMtARKUT8OZkDCUIQjKyNookAv4vcn4c10lFluhZHen6dGRrsutmQ9qz
# sIzV6Q3d9gEgzpkxYz0IGhizgZtPxpMQBvwHgfqL2vmCSfdibqFT+hKUGIUukpHq
# aGxEMrJmoecYpJpkUe8xggQ3MIIEMwIBATCBhjByMQswCQYDVQQGEwJVUzEVMBMG
# A1UEChMMRGlnaUNlcnQgSW5jMRkwFwYDVQQLExB3d3cuZGlnaWNlcnQuY29tMTEw
# LwYDVQQDEyhEaWdpQ2VydCBTSEEyIEFzc3VyZWQgSUQgQ29kZSBTaWduaW5nIENB
# AhAB7uHu9sSCKgqVgJoW5mKYMAkGBSsOAwIaBQCgeDAYBgorBgEEAYI3AgEMMQow
# CKACgAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3AgEEMBwGCisGAQQBgjcC
# AQsxDjAMBgorBgEEAYI3AgEVMCMGCSqGSIb3DQEJBDEWBBSLywPhDGo9oIp5FSjR
# cC+O+7Gx5jANBgkqhkiG9w0BAQEFAASCAQCrHaYAhO0FZfClrmXqK+QbuTyfcfrZ
# fU+uKBQoCA+1vqCgS7a86iQjD4MW4eFgqUggTAZ4plPJpZ/YLA2rwBI5ROkaIkuH
# /QCHyFKzaFwTgEQghlkBU4I8TSFXATlJEtLEk0WXSEbqZfcoXD9N0hR66Lo6tVal
# jwNzVNQpo1EAikq3DpfHDLbQubpravwkViez2xqDTCZ/2jE4k9kXUKSIIPf4sCO3
# eyinEGeSvnty/gA5/KdWfapVso4HHZYmuNZAMnSVmUE5bhvbG6rr53/msCxpH5lo
# ZCQoRxoj0i4CunZt/HsROReQwnj13nf1IaBCEdANW2aTTfzL4UGkeznKoYICCzCC
# AgcGCSqGSIb3DQEJBjGCAfgwggH0AgEBMHIwXjELMAkGA1UEBhMCVVMxHTAbBgNV
# BAoTFFN5bWFudGVjIENvcnBvcmF0aW9uMTAwLgYDVQQDEydTeW1hbnRlYyBUaW1l
# IFN0YW1waW5nIFNlcnZpY2VzIENBIC0gRzICEA7P9DjI/r81bgTYapgbGlAwCQYF
# Kw4DAhoFAKBdMBgGCSqGSIb3DQEJAzELBgkqhkiG9w0BBwEwHAYJKoZIhvcNAQkF
# MQ8XDTE5MDMyOTA5MzkxNVowIwYJKoZIhvcNAQkEMRYEFApWKQsb/rjLyovUH3iw
# TYKgOs31MA0GCSqGSIb3DQEBAQUABIIBAHp2spL71EQ/Bwrm22JnhN5FB56hsRUu
# TcNg8fnQCKaEBX+3W8ISPtYvzkVRTB3sm1Q55tgsGKnSl6B/xeOluMlxNp4e1B3L
# biflBms2j8840LMxficJegzhjUlSpZ68JmCAKZ6XFulXW1TcvsOa0XSVKf6fplyu
# 1dgn7wBlkFgjFuapNSmAh6mNTIePLDQlXPiAaMOAltCXpadBcvxHpdTC2Z4AXFue
# iZvYxEO6gFggKpjE+9cadyrNqYsCcwggt+jjse1ukOALXzgip0dQepTrm8toOZfw
# qH5YW8akF2Qx7Xr0Bk8YQC9Zlorut8kXhgtZqOeBzvDhSt1+qbQp2hE=
# SIG # End signature block
