Connect-MicrosoftTeams

#CSV INput for new Teams
$csvinput = Import-CSV -Path "" -Delimiter ";"

#Check API Delay
$delay = 10
#Max Loop for Check
$maxtry = 720
#PNP Client ID
$pnpclientid = ''
#PNP list for Documents
$pnplist = 'Dokumente'
#Temporary Admin for new Teams like m.mustermann-adm@example.com
$tempadmin = ''
#Company Sharepoint Site URL like https://example.sharepoint.com'
$sharepoint = ''

$AllowAddRemoveApps = $false
$AllowCreateUpdateChannels = $false
$AllowCreateUpdateRemoveConnectors = $false
$AllowCreateUpdateRemoveTabs = $false
$AllowDeleteChannels = $false
$AllowGuestCreateUpdateChannels = $false
$AllowGuestDeleteChannels = $false

foreach($zeile in $csvinput)
{
    $type = $zeile.name.Split("-")
	$globalowners = $zeile.besitzer.Split(",")
	$globalusers = $zeile.mitglieder.Split(",")
	$channelowners = $zeile.ownerstrictlyconfidentialchannel.Split(",")
	$channelusers = $zeile.mitgliederstrictlyconfidentialchannel.Split(",")
	
	$group = New-Team -MailNickname $zeile.name -displayname $zeile.name -Visibility $zeile.verfuegbarkeit -owner $tempadmin -AllowAddRemoveApps $AllowAddRemoveApps -AllowCreateUpdateChannels $AllowCreateUpdateChannels -AllowCreateUpdateRemoveConnectors $AllowCreateUpdateRemoveConnectors -AllowCreateUpdateRemoveTabs $AllowCreateUpdateRemoveTabs -AllowDeleteChannels $AllowDeleteChannels -AllowGuestCreateUpdateChannels $AllowGuestCreateUpdateChannels -AllowGuestDeleteChannels $AllowGuestDeleteChannels
    Start-Sleep -Seconds 10.0
	
    Connect-PnPOnline -Url $sharepoint -Interactive -ClientId $pnpclientid
    $SiteURL = Get-PnPMicrosoft365Group -Identity $group.groupid -IncludeSiteUrl | Select-Object -ExpandProperty SiteUrl
    #Disconnect-PnPOnline
    
    $tryloop = 0
    do
    {
	    try 
	    {
            Connect-PnPOnline -Url $SiteURL -Interactive -ClientId $pnpclientid
            break
        } 
	    catch 
	    {
            Write-Verbose "Error $tryloop"
		    $tryloop++	
        }
        Start-Sleep -Seconds $Delay

    }
    while($tryloop -le $maxtry)

    for ($i = 0; $i -lt $globalowners.Length; $i++) 
	{
		Add-TeamUser -GroupId $group.GroupId -User $globalowners[$i] -Role 'owner'
	}

    if($globalusers.length -gt 0)
	{
		for ($i = 0; $i -lt $globalusers.Length; $i++) 
		{	
            if($globalusers[$i])		
            {	
                Add-TeamUser -GroupId $group.GroupId -User $globalusers[$i] -Role member
            }
		}
	}

    if($type[1] -eq 'internal' -or $type[1] -eq 'customer' -or $type[1] -eq 'project')
    {
        
        $tryloop = 0
        do
        {
            try 
            {
                Add-PnPField -List $pnplist -DisplayName "confidentiality" -InternalName "confidentiality" -Type Choice -AddToDefaultView -Required -Choices "low","medium","high"
                break
            } 
            catch 
            {
                Write-Verbose "Error $tryloop"
                $tryloop++	
            }
            Start-Sleep -Seconds $Delay
    
        }
        while($tryloop -le $maxtry)
        
        Add-PnPField -List $pnplist -DisplayName "availability" -InternalName "availability" -Type Choice -AddToDefaultView -Required -Choices "low","medium","high"
        Add-PnPField -List $pnplist -DisplayName "Integrity" -InternalName "Integrity" -Type Choice -AddToDefaultView -Required -Choices "low","medium"
        
        Set-PnPDefaultColumnValues -List $pnplist -Field confidentiality -Value "medium"
        Set-PnPDefaultColumnValues -List $pnplist -Field availability -Value "medium"
        Set-PnPDefaultColumnValues -List $pnplist -Field Integrity -Value "medium"
        
        if($type[1] -eq 'customer')
        {
            $tryloop = 0
            do
            {
            try 
                {
                    Add-PnPFolder -Name Angebote -Folder "Freigegebene Dokumente/General"
                    break
                } 
                catch 
                {
                    Write-Verbose "Error $tryloop"
                    $tryloop++	
                }
                Start-Sleep -Seconds $Delay
    
            }
            while($tryloop -le $maxtry)            
            
            Add-PnPFolder -Name Auftraege_Projekte -Folder "Freigegebene Dokumente/General"
            Add-PnPFolder -Name Vertraege -Folder "Freigegebene Dokumente/General"
            Add-PnPFolder -Name Technik -Folder "Freigegebene Dokumente/General"
            Add-PnPFolder -Name Allgemein -Folder "Freigegebene Dokumente/General"
            Add-PnPFolder -Name Allgemein -Folder "Freigegebene Dokumente/General/Technik"
            Add-PnPFolder -Name Dokumentation -Folder "Freigegebene Dokumente/General/Technik"
            Add-PnPFolder -Name Zeichnungen -Folder "Freigegebene Dokumente/General/Technik"
            Add-PnPFolder -Name Konfig_Dateien -Folder "Freigegebene Dokumente/General/Technik"
            Add-PnPFolder -Name Bilder -Folder "Freigegebene Dokumente/General/Technik"
            Add-PnPFolder -Name Zertifikate -Folder "Freigegebene Dokumente/General/Technik"
            Add-PnPFolder -Name Netzwerk -Folder "Freigegebene Dokumente/General/Technik"
            Add-PnPFolder -Name Software -Folder "Freigegebene Dokumente/General/Technik"
            Add-PnPFolder -Name Logs -Folder "Freigegebene Dokumente/General/Technik"
            Add-PnPFolder -Name Archiv -Folder "Freigegebene Dokumente/General/Technik"
        }
        elseif($type[1] -eq 'project')
        {
            $tryloop = 0
            do
            {
            try 
                {
                    Add-PnPFolder -Name Angebote -Folder "Freigegebene Dokumente/General"
                    break
                } 
                catch 
                {
                    Write-Verbose "Error $tryloop"
                    $tryloop++	
                }
                Start-Sleep -Seconds $Delay
    
            }
            while($tryloop -le $maxtry) 

            Add-PnPFolder -Name Projektplan -Folder "Freigegebene Dokumente/General"
            Add-PnPFolder -Name Vertrieb -Folder "Freigegebene Dokumente/General"
            Add-PnPFolder -Name Technik -Folder "Freigegebene Dokumente/General"
            Add-PnPFolder -Name Allgemein -Folder "Freigegebene Dokumente/General"
        }
    }
    if($type[1] -eq 'private')
    {
        $tryloop = 0
        do
        {
            try 
            {
                Add-PnPField -List $pnplist -DisplayName "confidentiality" -InternalName "confidentiality" -Type Choice -AddToDefaultView -Required -Choices "low","medium"
                break
            } 
            catch 
            {
                Write-Verbose "Error $tryloop"
                $tryloop++	
            }
            Start-Sleep -Seconds $Delay
    
        }
        while($tryloop -le $maxtry)
        
        Add-PnPField -List $pnplist -DisplayName "availability" -InternalName "availability" -Type Choice -AddToDefaultView -Required -Choices "low"
        Add-PnPField -List $pnplist -DisplayName "Integrity" -InternalName "Integrity" -Type Choice -AddToDefaultView -Required -Choices "low","medium"
        
        Set-PnPDefaultColumnValues -List $pnplist -Field confidentiality -Value "medium"
        Set-PnPDefaultColumnValues -List $pnplist -Field availability -Value "low"
        Set-PnPDefaultColumnValues -List $pnplist -Field Integrity -Value "medium"
    }
    if($type[1] -eq 'public')
    {
        $tryloop = 0
        do
        {
            try 
            {
                Add-PnPField -List $pnplist -DisplayName "confidentiality" -InternalName "confidentiality" -Type Choice -AddToDefaultView -Required -Choices "medium"
                break
            } 
            catch 
            {
                Write-Verbose "Error $tryloop"
                $tryloop++	
            }
            Start-Sleep -Seconds $Delay
    
        }
        while($tryloop -le $maxtry)

        Add-PnPField -List $pnplist -DisplayName "availability" -InternalName "availability" -Type Choice -AddToDefaultView -Required -Choices "low","medium"
        Add-PnPField -List $pnplist -DisplayName "Integrity" -InternalName "Integrity" -Type Choice -AddToDefaultView -Required -Choices "low","medium"
        
        Set-PnPDefaultColumnValues -List $pnplist -Field confidentiality -Value "medium"
        Set-PnPDefaultColumnValues -List $pnplist -Field availability -Value "medium"
        Set-PnPDefaultColumnValues -List $pnplist -Field Integrity -Value "medium"
    }

    #Disconnect-PnPOnline

    if($type[1] -eq 'internal' -or $type[1] -eq 'customer' -or $type[1] -eq 'project')
    {
        New-TeamChannel -GroupId $group.GroupId -DisplayName "top secret" -MembershipType Private -owner $tempadmin

        Start-Sleep -Seconds 10.0

        for ($i = 0; $i -lt $channelowners.Length; $i++) 
        {
            Add-TeamChannelUser -GroupId $group.GroupId -DisplayName "top secret" -User $channelowners[$i]
            Add-TeamChannelUser -GroupId $group.GroupId -DisplayName "top secret" -User $channelowners[$i] -Role owner
        }
        
        if($channelusers.length -gt 0)
        {
            for ($i = 0; $i -lt $channelusers.Length; $i++) 
            {
                if($channelusers[$i])		
                {	
                    Add-TeamChannelUser -GroupId $group.GroupId -DisplayName "top secret" -User $channelusers[$i]
                }
            }
        }

        #$channelsite = $zeile.name + '-strictlyconfidential'
        $channelURL = $SiteURL + '-topsecret'

        $tryloop = 0
        do
        {
            try 
            {
                Connect-PnPOnline -Url $channelURL -Interactive -ClientId $pnpclientid
                break
            } 
            catch 
            {
                Write-Verbose "Error $tryloop"
                $tryloop++	
            }
            Start-Sleep -Seconds $Delay
     
        }
        while($tryloop -le $maxtry)       

        $tryloop = 0
        do
        {
            try 
            {
                Add-PnPField -List $pnplist -DisplayName "confidentiality" -InternalName "confidentiality" -Type Choice -AddToDefaultView -Required -Choices "top secret"
                break
            } 
            catch 
            {
                Write-Verbose "Error $tryloop"
                $tryloop++	
            }
            Start-Sleep -Seconds $Delay
     
        }
        while($tryloop -le $maxtry)

        Add-PnPField -List $pnplist -DisplayName "availability" -InternalName "availability" -Type Choice -AddToDefaultView -Required -Choices "low","medium","high"
        Add-PnPField -List $pnplist -DisplayName "Integrity" -InternalName "Integrity" -Type Choice -AddToDefaultView -Required -Choices "low","medium"
        
        Set-PnPDefaultColumnValues -List $pnplist -Field confidentiality -Value "top secret"
        Set-PnPDefaultColumnValues -List $pnplist -Field availability -Value "medium"
        Set-PnPDefaultColumnValues -List $pnplist -Field Integrity -Value "medium"

        #Disconnect-PnPOnline

        Remove-TeamChannelUser -GroupId $group.GroupId -DisplayName "top secret" -User $tempadmin
        Remove-TeamUser -GroupId $group.GroupId -User $tempadmin
    }
}
Disconnect-PnPOnline