# Importujte modul Active Directory
# Import-Module ActiveDirectory

# Definujte proměnné
$backupFolders = @("C:\Slozka1", "C:\Slozka2")  # Zadejte složky, které chcete zálohovat
$backupDestination = "C:\Backup"  # Cílová složka pro zálohu
$emailTo = "administrator@ee.cz"  # E-mailová adresa příjemce
$smtpServer = "emailserver.ee.cz"  # SMTP server pro odesílání e-mailů
$smtpFrom = "nobody@ee.cz"  # E-mailová adresa odesílatele
$zipExePath = "C:\Program Files\7-Zip\7z.exe"  # Cesta k 7-Zip executable

# Funkce pro získání jmen stanic z AD
function Get-StationNamesFromAD {
    param (
        [string]$filter = "*"  # Filtr pro vyhledávání stanic v AD
    )
    # $computers = Get-ADComputer -Filter "Name -like '$filter'" | Select-Object -ExpandProperty Name
    # return $computers  # Vrácení seznamu jmen stanic
    return "ws14"
}

# Funkce pro vytvoření zálohy
function Backup-Folders {
    param (
        [string[]]$folders,  # Pole složek k zálohování
        [string]$destination  # Cílová složka pro zálohu
    )
    $tempBackupPath = Join-Path $destination "TempBackup"
    $tempFolder=$tempBackupPath
    if (Test-Path $tempBackupPath) {
        Remove-Item $tempBackupPath -Recurse -Force
    } else {
        Write-Host "Cesta $tempBackupPath neexistuje."
    }
    New-Item -ItemType Directory -Path $tempBackupPath
    foreach ($folder in $folders) {  # Pro každou složku v poli složek
        $folderName = Split-Path $folder -Leaf  # Získání názvu složky
        $destPath = Join-Path $tempBackupPath $folderName  # Vytvoření cesty k dočasné záložní složce
        Copy-Item -Path "$folder" -Destination "$destPath" -Recurse  # Kopírování složky do dočasné záložní složky
    }
    # Výpis obsahu dočasné záložní složky pro kontrolu
    Get-ChildItem -Path $tempBackupPath -Recurse
    return $tempBackupPath
}

# Funkce pro zazipování složek s heslem
function Zip-FoldersWithPassword {
    param (
        [string]$source,
        [string]$destination,
        [string]$zipFileName,
        [string]$password
    )
    $zipFilePath = Join-Path $destination $zipFileName
    if (Test-Path $zipFilePath) {
        Remove-Item $zipFilePath -Force
    }
    $tempFolder=$source
    # Připravte argumenty pro 7-Zip
   # $arguments = "a -tzip `"$zipFilePath`" `"$source\*`" -p$password"
   # $arguments = "a -tzip `"$zipFilePath`" `"C:\Backup\TempBackup\*`" -p$password"
    $arguments = "a -tzip `"$zipFilePath`" `"$tempFolder\*`" -p$password"
    
    Write-Host "Spouštím 7-Zip s argumenty: $arguments"
    Start-Process -FilePath $zipExePath -ArgumentList $arguments -Wait
    return $zipFilePath
}

# Funkce pro odeslání e-mailu
function Send-Email {
    param (
        [string]$to,  # E-mailová adresa příjemce
        [string]$subject,  # Předmět e-mailu
        [string]$body,  # Tělo e-mailu
        [string]$smtpServer,  # SMTP server pro odesílání e-mailů
        [string]$from  # E-mailová adresa odesílatele
    )
    $message = New-Object System.Net.Mail.MailMessage  # Vytvoření nového e-mailu
    $message.From = $from  # Nastavení odesílatele
    $message.To.Add($to)  # Přidání příjemce
    $message.Subject = $subject  # Nastavení předmětu
    $message.Body = $body  # Nastavení těla e-mailu

    $smtp = New-Object Net.Mail.SmtpClient($smtpServer)  # Vytvoření nového SMTP klienta
    try {
        $smtp.Send($message)
        Write-Host "E-mail byl úspěšně odeslán."
    } catch {
        Write-Host "Chyba při odesílání e-mailu: $_"
    }
}

# Funkce pro získání sériového čísla počítače
function Get-SerialNumber {
    $wmi = Get-WmiObject -Class Win32_BIOS
    return $wmi.SerialNumber
}

# Hlavní skript
$stationNames = Get-StationNamesFromAD -filter "*"  # Získání všech jmen stanic z AD
if ($stationNames -contains $env:COMPUTERNAME) {  # Pokud je aktuální stanice v seznamu jmen stanic
    $tempBackupPath = Backup-Folders -folders $backupFolders -destination $backupDestination  # Vytvoření zálohy složek
    $zipFileName = "Backup_$(Get-Date -Format 'yyyyMMddHHmmss').zip"
    $encryptionKey = [System.Guid]::NewGuid().ToString("N").Substring(0, 32)  # Generování náhodného klíče pro šifrování (32 znaků)
    $zipFilePath = Zip-FoldersWithPassword -source "C:\Backup\TempBackup" -destination $backupDestination -zipFileName $zipFileName -password $encryptionKey  # Zazipování složek s heslem
    Remove-Item $tempBackupPath -Recurse -Force  # Odstranění dočasné záložní složky
   $emailBody = "The backup has been successfully created and encrypted. The password for the zip file is: $encryptionKey"
    Send-Email -to $emailTo -subject "Backup Completed" -body  $emailBody -smtpServer $smtpServer -from $smtpFrom  # Odeslání e-mailu
}