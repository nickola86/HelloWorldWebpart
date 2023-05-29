. .\modules\module-login.ps1

#declare
$spBaseUrl = "https://7mcbww.sharepoint.com/"
$siteUrl = $spBaseUrl + "sites/dev"


try{
    Do-SPLogin($siteUrl) 1>$null 2>$null
}catch{
    write-host "Errore bloccante durante la connessione al sito " + $siteUrl
    write-host "Internal error: " $_.Exception.Message
    Exit -3    
}

Add-PnPField -DisplayName "Cognome" -InternalName "Cognome" -Type Text 
Add-PnPField -DisplayName "Nome" -InternalName "Nome" -Type Text 
Add-PnPField -DisplayName "Luogo di nascita" -InternalName "LuogoDiNascita" -Type Text 
Add-PnPField -DisplayName "Genere" -InternalName "Genere" -Type Text 
Add-PnPField -DisplayName "Data di nascita" -InternalName "DataDiNascita" -Type DateTime 

Add-PnPContentType -Name "Utente" -Description "Custom Content Types" -Group "Custom Content Types" -ParentContentType (Get-PnPContentType Item)

Add-PnPFieldToContentType -Field "Cognome" -ContentType "Utente"
Add-PnPFieldToContentType -Field "Nome" -ContentType "Utente"
Add-PnPFieldToContentType -Field "LuogoDiNascita" -ContentType "Utente"
Add-PnPFieldToContentType -Field "Genere" -ContentType "Utente"
Add-PnPFieldToContentType -Field "DataDiNascita" -ContentType "Utente"

New-PnPList -Title "Utenti" -Template GenericList -EnableContentTypes
Add-PnPContentTypeToList -List "Utenti" -ContentType "Utente"
Add-PnPListItem -List "Utenti" -Values @{"Cognome"="Di Trani"; "Nome"="Nicola"; "LuogoDiNascita"="Andria (BT)"; "Genere"="Maschile";}
Add-PnPListItem -List "Utenti" -Values @{"Cognome"="Rossi"; "Nome"="Giorgio"; "LuogoDiNascita"="Roma (RM)"; "Genere"="Maschile";}
Add-PnPListItem -List "Utenti" -Values @{"Cognome"="Della Valle"; "Nome"="Valeria"; "LuogoDiNascita"="Savona (SV)"; "Genere"="Femminile";}
