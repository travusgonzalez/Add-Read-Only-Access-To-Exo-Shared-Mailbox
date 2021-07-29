######################################################################################
#Check for Basic Authentication Reg Entry and create/modify if not present
######################################################################################
#https://docs.microsoft.com/en-us/powershell/exchange/exchange-online-powershell-v2
    
    $regPath = "HKLM:\Software\Policies\Microsoft\Windows\WinRM\Client\"
    #key entry
    $Kname = "AllowBasic"
    $Kvalue = "1"
    Write-Host "Checking Basic Authentication Key..." -ForegroundColor green
    #check if reg path doesn't exist yet
    IF(!(Test-Path $regPath)){
        #create registry path
            Write-Output 'The basic auth key does not exist, creating key...'
            New-Item -Path $regPath -Force
        #add reg entry
            New-ItemProperty -Path $regPath -Name $Kname -Value $Kvalue -PropertyType DWORD -Force
    }
    ELSE {
        # if reg path exists, check if registry key exists
        Try {
            Get-ItemPropertyValue -Path $regPath -Name $Kname
        }
        Catch {
            Write-Output 'The basic auth key does not exist, creating key...'
            New-ItemProperty -Path $regPath -Name $Kname -Value $Kvalue -PropertyType DWORD -Force
        }
        Finally {
            $startValue = Get-ItemPropertyValue -Path $regPath -Name $Kname
            if ($startValue -eq '1') {
                Write-Output 'Basic Auth is enabled'
            } else {
                Write-Output 'Basic Auth is disabled/not configured. Modifying key to enable.'
                Set-ItemProperty -Path $regPath -Name $Kname -Value $Kvalue
            }
        }
    }