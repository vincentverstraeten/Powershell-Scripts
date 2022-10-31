#=============================================================================================================================
# Script Name: Backup Outlook Signature to onedrive
# Description: Symlink Outlook Signature to onedrive   New-Item -ItemType SymbolicLink -Path $LocalOutlookSignature  -Target $SignatureOnedrive -Force
# Notes: Vincent Verstraeten 08/2022     New-Item -Path $SignatureOnedrive -ItemType "directory" -Force

#=============================================================================================================================

# Main script
# get local user folders
$env:APPDATA

$LocalOutlookSignature = Join-Path $env:APPDATA -ChildPath "Microsoft/Signatures"
$LocalOutlookBackupSignature = Join-Path $env:APPDATA -ChildPath "Microsoft/Signatures_backup"
$OneDrivePath = [Environment]::GetEnvironmentVariable("ONEDRIVE",[EnvironmentVariableTarget]::User)   
$SignatureOnedrive = join-path -Path $OneDrivePath -ChildPath "Signatures"



if ((Test-Path -Path $LocalOutlookSignature -ErrorAction SilentlyContinue) -or ((Get-Item -Path $LocalOutlookSignature -Force -ErrorAction SilentlyContinue).LinkType -eq "Junction"))  {
    Copy-Item -Path $LocalOutlookSignature -Destination $LocalOutlookBackupSignature -recurse -Force
    Remove-Item -LiteralPath $LocalOutlookSignature -Force -Recurse
        if(!(Test-Path -Path $SignatureOnedrive)){
                mkdir $SignatureOnedrive
                # Make onedrive map
         }
        else{
                # Onedrive map is already there
          }
    cmd /c mklink /J $LocalOutlookSignature $SignatureOnedrive #Powershell command needs admin, only in preview windows(better use cmd here)
    Copy-Item -Path "$LocalOutlookBackupSignature\*" -Destination $LocalOutlookSignature -recurse -Force
    Remove-Item -LiteralPath $LocalOutlookBackupSignature -Force -Recurse
    # Copy over local signature data to onedrive
}

#Check if onedrive signature map is present but local microsoft signature outlook is not there. Then copy Onedrive Signature to outlook signature.

if (!(Test-Path $LocalOutlookSignature)) {
    Copy-Item -Path $SignatureOnedrive -Destination $LocalOutlookBackupSignature -recurse -Force
    cmd /c mklink /J $LocalOutlookSignature $SignatureOnedrive
    Copy-Item -Path $LocalOutlookBackupSignature\* -Destination $SignatureOnedrive  -recurse -Force
    Remove-Item -LiteralPath $LocalOutlookBackupSignature -Force -Recurse
    # Copy onedrive to outlook signature
}

if (!((Get-Item -Path $LocalOutlookSignature  -Force).LinkType -eq "Junction")) {
    Remove-Item -LiteralPath $LocalOutlookSignature -Force -Recurse
    cmd /c mklink /J $LocalOutlookSignature $SignatureOnedrive
}