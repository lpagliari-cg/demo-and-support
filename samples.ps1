# RACOCCOLTA DI QUALCHE SCRIPT PnP POWERSHELL DI ESEMPIO

####### PRE REQUISITI #############################################################
# STEP 1: installare il modulo di PnP PowerShell
Install-Module PnP.PowerShell -Scope CurrentUser
Update-Module PnP.PowerShell -Scope CurrentUser
 
# STEP 2: dare i permessi all'applicazione PnP PowerShell su tenant 
# NOTA: va fatto una sola volta per ogni tenant
# Register-PnPEntraIDAppForInteractiveLogin -ApplicationName "PnP Rocks" -Tenant [yourtenant].onmicrosoft.com -Interactive
Register-PnPEntraIDAppForInteractiveLogin -ApplicationName "PnP Rocks" -Tenant m365x13866366.onmicrosoft.com -Interactive
# NOTA: salvare il clientID restituito dal comando precedente
 
########### CREAZIONE NUOVO SITO + LIBRARY ########################################
# Connessione all'admin center di SharePoint
$clientId = "ba9e89f5-ed71-4d49-ab2a-d2fcee5651b1"
$tenant = "m365x13866366"
Connect-PnPOnline -Url "https://m365x13866366-admin.sharepoint.com" -ClientId $clientId -Interactive
# Creazione di un nuovo sito del team slegato da M365 group con lingua default italiano
$newSiteUrl = "https://m365x13866366.sharepoint.com/sites/SitoScript"
New-PnPSite -Type TeamSiteWithoutMicrosoft365Group -Title "Sito Script" -Url $newSiteUrl -Lcid 1040
 
# Connessione al nuovo sito
Connect-PnPOnline -Url $newSiteUrl -ClientId $clientId -Interactive
# Creazione di una nuova libreria documenti chiamata "Fatture"
New-PnPList -Title "Fatture" -Template DocumentLibrary
# Aggiungo la libreria alla navigazione del sito
$libraryName = "Fatture"
Add-PnPNavigationNode -Title $libraryName -Url "$newSiteURL/$libraryName" -Location "QuickLaunch"
# Interrompo l'ereditarietà dei permessi sulla nuova libreria
Set-PnPList -Identity "Fatture" -BreakRoleInheritance

########## CREAZIONE NUOVO SITO + SECURITY GROUPS + MODIFICA INTERFACCIA ###############
# Tenant variables
$tenant = "DOMAIN"
$clientId = "CLIENT ID DA APP REGISTRATION"
# Site variables
$siteName = "SITE NAME"
$newSiteURL = "https://$tenant.sharepoint.com/sites/SITENAMENOSPACE"
 
## NEW SITE CREATION ##
# Tenant admin center connection
Connect-PnPOnline -Url "https://$tenant.sharepoint.com" -ClientId $clientId -Interactive
# New site without microsoft 365 group and italian default language
New-PnPSite -Type TeamSiteWithoutMicrosoft365Group -Title $siteName -Url $newSiteURL -Lcid 1040 
# New site connection
Connect-PnPOnline -Url $newSiteURL -ClientId $clientId -Interactive
# Create a new AAD security group named as the site and add it to site visitors SharePoint group 
$displayName = "sg-$siteName"
$newADGroup = New-PnPAzureADGroup -DisplayName $displayName -MailNickname $siteName -Description "Temp" -IsSecurityEnabled
$visitorGroup = Get-PnPGroup -AssociatedVisitorGroup
# NOTA: se partite da security group già esistenti vi togliete il problema di dover aspettare il provisioning
$condition = $false
do{
    try {
        # unico comando effettivo, il resto è solo aspettare il provisioning in ciclo
        Add-PnPGroupMember -Group $visitorGroup -LoginName ("c:0t.c|tenant|" + $newADGroup.Id)
        $condition = $true
    }
    catch {
        # workaround che non tiene conto di altri errori, se va in loop terminare lo script con CTRL+C su terminale
        Write-Output "Waiting for $siteName security group provisioning..." 
        Start-Sleep -Seconds 15
    }
}
until($condition)
# Esempio di "TEMPLATE MANUALE", andando a modificare le componenti del sito pezzo per pezzo
# Remove all navigation nodes except Home
$navNodes = Get-PnPNavigationNode
foreach ($node in $navNodes) {
    if ($node.Title -ne "Home.aspx") {
        Remove-PnPNavigationNode -Identity $node.Id -Force
    }
}
# Remove the default document library for italian site
# JUST AN EXAMPLE - NOT A BEST PRACTICE
Remove-PnPList -Identity "Documenti condivisi" -Force
# Get the home page and remove all components except Activities
Add-PnPPageWebPart -Page "Home" -DefaultWebPartType SiteActivity
Set-PnPPage -Identity "Home" -Publish
Write-Output "Homepage cleaning done." 
