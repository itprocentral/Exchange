#Defaul variables for the script
$vhost = hostname
$vSource = "\\SERVER\SHARE\SUBFOLDER"
$vinfo = import-csv ($vSource + "\customer.info")
$vserver = import-csv ($vSource + "\server.info")

clear
Write-host
Write-host "Exchange Deployment Tool" -ForegroundColor Yellow
Write-Host
write-host ".:. OS Settings" -foregroundcolor green
Write-host "1 - Network adjustments"
Write-host "2 - Time Zone"
Write-Host "3 - Pagefile configuration"
Write-host "4 - .Net fix (temporary solution)" 
Write-Host
write-host ".:.Exchange Server 2016" -foregroundcolor green
Write-host "10 - Exchange 2016 - Installation files.. <Copy Only>"
write-host "11 - Exchange 2016 - OS requirements with restart"
write-host "12 - UCMA" 
write-host "13 - Exchange 2016 - Deployment"
write-host
write-host ".:. Exchange Server Settings (Requires Exchange Management Shell)" -foregroundcolor green
Write-host "30 - Configure Web Services"
write-host "31 - Autodiscover"
write-host "32 - License"
write-host "33 - Exchange Certificate"
write-host "34 - Outlook Settings"
write-host "35 - Message Tracking Settings"
write-host "36 - OWA Settings"
write-host "37 - Managed Folder Settings"
write-host "38 - Transport Settings"
write-host "39 - Configure Hybrid Cloud Settings"
write-host 
write-host
Write-Host 0 - Operator or Exit
write-host
$opt = Read-Host -Prompt "Select your option"

#Preparing the environment..
$vPath = "C:\Temp\"
If (Test-Path $vPath) {} Else{New-Item -Path C:\Temp -ItemType Directory}

If ($opt -eq 1)
    {
        Write-Host "Current Network Adapters in this server.."
        $vStatus = 0
        get-netadapter MAPI -ErrorAction SilentlyContinue -ErrorVariable erNet
        If ($ernet.count -ne 0) { $vStatus = 1}
        get-netadapter Replication01 -ErrorAction SilentlyContinue -ErrorVariable erNet
        If ($ernet.count -ne 0) { $vStatus = 1}
        get-netadapter Replication02 -ErrorAction SilentlyContinue -ErrorVariable erNet
        If ($ernet.count -ne 0) { $vStatus = 1}
        get-netadapter SMTP -ErrorAction SilentlyContinue -ErrorVariable erNet
        If ($ernet.count -ne 0) { $vStatus = 1}

        If ($vStatus -eq 1) { Write-Host "The adapters are not in compliance with the Exchange Design"}
	        Else {
                #Defining IPs for replication networks
                New-NetIPAddress -InterfaceAlias Replication01 -IPAddress ($vserver | Where-Object {$_.Server -eQ $vhost}).NICRep01 -PrefixLength 24
                New-NetIPAddress -InterfaceAlias Replication02 -IPAddress ($vserver | Where-Object {$_.Server -eQ $vhost}).NICRep02 -PrefixLength 24
                New-NetIPAddress -InterfaceAlias SMTP -IPAddress ($vserver | Where-Object {$_.Server -eQ $vhost}).SMTP -PrefixLength 24
                Write-Host "Configuring Network Adapters..." 
		        Get-netadapter Replication01 | Set-DnsClient -RegisterThisConnectionsAddress $false
		        Get-netadapter Replication02 | Set-DNSClient -RegisterThisConnectionsAddress $False
                Get-netadapter SMTP | Set-DNSClient -RegisterThisConnectionsAddress $False
		        wmic /interactive:off nicconfig where tcpipnetbiosoptions=0 call SetTcpipNetbios 2
                Disable-NetAdapterBinding Replication01 -ComponentID ms_server
                Disable-NetAdapterBinding Replication01 -ComponentID ms_msclient
                Disable-NetAdapterBinding Replication01 -ComponentID ms_pacer
                Disable-NetAdapterBinding Replication02 -ComponentID ms_server
                Disable-NetAdapterBinding Replication02 -ComponentID ms_msclient
                Disable-NetAdapterBinding Replication02 -ComponentID ms_pacer
		        write-host "all good!"
	        }
    }

If ($opt -eq 2)
    {
        write-host "Current Time Zone on server: " ($vserver | Where-Object {$_.Server -eq $vhost}).TimeZone
        tzutil /s ($vserver | Where-Object {$_.Server -eq $vhost}).TimeZone
        Write-Host "New Time Zone configured based on the design:" 
        tzutil /g
    }

If ($opt -eq 3)
    {
        Write-host "Configuring the pagefile based on the design.. restart required afterwards"
        Set-CimInstance -Query "SELECT * FROM Win32_computersystem" -Property @{AutomaticManagedPageFile="False"}
        Set-CimInstance -Query "SELECT * FROM Win32_PageFileSetting" -Property @{InitialSize=32768;MaximumSize=32778}
    }

If ($opt -eq 4)
    {
        write-host "Configuring .Net Fix..."
        C:\windows\Microsoft.net\Framework64\v4.0.30319\ngen.exe update
    }

If ($opt -eq 10)
    {
    write-host
    write-host "Exchange 2016 - Installation Files.. copying from the source (it may take a while..) go for a Tim's"
    If (Test-Path "C:\temp\Deployment") {} Else{New-Item -Path C:\Temp\Deployment -ItemType Directory}
    Copy-Item ($vSource + "\EXDeployment\*") C:\Temp\Deployment\ -Recurse -Force
    write-host
    }


If ($opt -eq 11)
    {
    write-host
    write-host "Exchange Requirements with restart"
    Install-WindowsFeature NET-Framework-45-Features, RPC-over-HTTP-proxy, RSAT-Clustering, RSAT-Clustering-CmdInterface, RSAT-Clustering-Mgmt, RSAT-Clustering-PowerShell, Web-Mgmt-Console, WAS-Process-Model, Web-Asp-Net45, Web-Basic-Auth, Web-Client-Auth, Web-Digest-Auth, Web-Dir-Browsing, Web-Dyn-Compression, Web-Http-Errors, Web-Http-Logging, Web-Http-Redirect, Web-Http-Tracing, Web-ISAPI-Ext, Web-ISAPI-Filter, Web-Lgcy-Mgmt-Console, Web-Metabase, Web-Mgmt-Console, Web-Mgmt-Service, Web-Net-Ext45, Web-Request-Monitor, Web-Server, Web-Stat-Compression, Web-Static-Content, Web-Windows-Auth, Web-WMI, Windows-Identity-Foundation, RSAT-ADDS, Telnet-Client
    Restart-Computer -Force    
    write-host
    }
If ($opt -eq 12)
    {
    write-host
    write-host "Installing UCMA.."
	C:\Temp\Deployment\UCMARuntimeSetup.exe /q  
    Restart-Computer -Force
    write-host
    }

If ($opt -eq 13)
    {
    write-host
    write-host "Installing ... - Exchange Server 2016"
    Mount-DiskImage "C:\Temp\Deployment\ExchangeServer2016-X64-CU4.iso"
    $vinstall = (Get-Volume -FileSystemLabel "ExchangeServer2016-X64-CU4").DriveLetter
    $vinstall = $vinstall +":\Setup.exe"
    write-host $vinstall
    [System.Diagnostics.Process]::Start("$vinstall"," /Mode:Install /Roles:Mailbox /MDBName:Temp01 /IAcceptExchangeServerLicenseTerms /DisableAMFiltering /InstallWindowsComponents /CustomerFeedbackEnabled:False")
    }

#Exchange Settings...

If ($opt -eq 30)
    {
        Set-ECPVirtualDirectory "$vhost\ECP (Default Web Site)" -InternalURL ("https://" + ($vinfo).URL + "/ecp") -ExternalURL ("https://" + ($vinfo).URL + "/ecp")
        Set-WebServicesVirtualDirectory "$vhost\EWS (Default Web Site)" -InternalURL ("https://" + ($vinfo).URL + "/EWS/Exchange.asmx") -ExternalURL ("https://" + ($vinfo).URL + "/EWS/Exchange.asmx")
        Set-ActiveSyncVirtualDirectory "$vhost\Microsoft-Server-ActiveSync (Default Web Site)" -InternalURL ("https://" + ($vinfo).URL + "/Microsoft-Server-ActiveSync") -ExternalURL ("https://" + ($vinfo).URL + "/Microsoft-Server-ActiveSync")
        Set-OABVirtualDirectory "$vhost\OAB (Default Web Site)" -InternalURL ("https://" + ($vinfo).URL + "/OAB") -ExternalURL ("https://" + ($vinfo).URL + "/OAB")
        Set-OWAVirtualDirectory "$vhost\OWA (Default Web Site)" -InternalURL ("https://" + ($vinfo).URL + "/OWA") -ExternalURL ("https://" + ($vinfo).URL + "/OWA")
        #Set-PowerShellVirtualDirectory "$vhost\PowerShell (Default Web Site)" -InternalURL ("https://" + ($vinfo).URL + "/powershell") -ExternalURL ("https://" + ($vinfo).URL + "/powershell")
        #Ouput
        Write-Host "New settings.."
        Get-EcpVirtualDirectory "$vhost\ecp (Default Web Site)" | fl Identity,InternalURL,ExternalURL
        Get-WebServicesVirtualDirectory "$vhost\ews (Default Web Site)" | fl Identity,InternalURL,ExternalURL
        Get-ActiveSyncVirtualDirectory "$vhost\Microsoft-Server-ActiveSync (Default Web Site)" | fl Identity,InternalURL,ExternalURL
        Get-OABVirtualDirectory "$vhost\oab (Default Web Site)" | fl Identity,InternalURL,ExternalURL
        Get-OWAVirtualDirectory "$vhost\owa (Default Web Site)" | fl Identity,InternalURL,ExternalURL
        #Get-PowerShellVirtualDirectory "$vhost\PowerShell (Default Web Site)" | fl Identity,InternalURL,ExternalURL
    }

If ($opt -eq 31)
    {
        write-host
        write-host "Autodiscover.. " $vhost
        Set-ClientAccessService -Identity $vhost -AutoDiscoverServiceInternalUri ("https://" + ($vinfo).autodiscover +"/Autodiscover/Autodiscover.xml")
        Get-ClientAccessService $vhost | Select Name,AutoDiscoverServiceInternalUri        
    }
If ($opt -eq 32)
    {
    write-host
    write-host "Applying License on the local server.. " $vhost
    set-exchangeserver $vhost -ProductKey ($vinfo).License
    write-host "- Settings configured based on the design.."
    Get-ExchangeServer | Select 
    write-host
    }

If ($opt -eq 33)
    {
    write-host
    write-host "Certificate.." $vhost
    Copy-Item -Path ($vSource + "\cert.pfx") -Destination C:\Temp\cert.pfx
    #Import-ExchangeCertificate -Server $vhost -FileName ($vSource + "\cert.pfx") -Password (ConvertTo-SecureString -String "m@nager171" -AsPlainText -Force)
    Import-ExchangeCertificate -Server $vhost -FileName "C:\temp\cert.pfx" -Password (ConvertTo-SecureString -String "m@nager171" -AsPlainText -Force)
    Get-ExchangeCertificate | where-object {$_.RootCAType -eq 'ThirdParty'} | Enable-ExchangeCertificate -Services IIS
    write-host "- Settings configured based on the design.."
    Get-ExchangeCertificate
    write-host
    }
If ($opt -eq 34)
    {
    write-host
    write-host "Outlook.." $vhost
    Set-OutlookAnywhere -Identity "$vhost\rpc (Default Web Site)" -InternalHostname ($vinfo).url -ExternalHostname ($vinfo).url -ExternalClientsRequireSsl $True -ExternalClientAuthenticationMethod 'NTLM' -InternalClientsRequireSsl $True
    write-host "- Settings configured based on the design.."
    Get-OutlookAnywhere -Identity "$vhost\rpc (Default Web Site)" | fl Identity,InternalHostname,ExternalHostName
    write-host
    }

If ($opt -eq 35)
    {
    write-host
    write-host "Message Tracking.." $vhost
    Set-TransportService $vhost -MessageTrackingLogEnabled $True -MessageTrackingLogMaxFileSize 5MB -MessageTrackingLogMaxDirectorySize 30GB -MessageTrackingLogSubjectLoggingEnabled $True -MessageTrackingLogMaxAge 60.00:00:00
    write-host "- Settings configured based on the design.." 
    Get-TransportService $vhost | fl MessageTrackingLogEnabled,MessageTrackingLogMaxFileSize,MessageTrackingLogMaxDirectorySize,MessageTrackingLogSubjectLoggingEnabled,MessageTrackingLogMaxAge
    write-host
    }

If ($opt -eq 36)
    {
    write-host
    write-host "Configuring OWA.." $vhost
    Set-OWAVirtualDirectory "$vhost\owa (Default Web Site)" -LogonPagePublicPrivateSelectionEnabled $True
    write-host "- Settings configured based on the design.." 
    Get-OWAVirtualDirectory "$vhost\owa (Default Web Site)" | Select Identity,LogonPagePublic*
    write-host
    }

If ($opt -eq 37)
    {
    write-host
    write-host "Managed Folder settings on" $vhost
    Set-MailboxServer $vhost -ManagedFolderAssistantSchedule @("Monday.1:00 AM-Monday.4:30 AM","Tuesday.1:00 AM-Tuesday.4:30 AM","Wednesday.1:00 AM-Wednesday.4:30 AM","Thursday.1:00 AM-Thursday.4:30 AM","Friday.1:00 AM-Friday.4:30 AM","Saturday.1:00 AM-Saturday.6:00 AM","Sunday.1:00 AM-Sunday.6:30 AM")
    write-host
    }

If ($opt -eq 38)
    {
    write-host
    write-host "Transport settings on" $vhost
    Set-TransportService $vhost -MaxOutboundConnections 50 -MaxPerDomainOutboundConnections 50
    }

If ($opt -eq 39)
    {
    write-host
    write-host "Configuring Hybrid Cloud Settings.. " $vhost
    Set-WebServicesVirtualDirectory "$vhost\EWS (Default Web Site)" -MRSProxyEnabled $True
    }

If ($opt -eq 0)
    {
    write-host
    write-host "Goodbye! May the Force be with you"
    write-host
    }
