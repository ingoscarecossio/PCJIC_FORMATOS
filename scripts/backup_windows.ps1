$ErrorActionPreference = "Stop"
$stamp = Get-Date -Format "yyyyMMdd_HHmmss"
$dest = "backup_fdgc_$stamp.zip"
Compress-Archive -Path app_data -DestinationPath $dest -Force
Write-Host "Respaldo creado: $dest"
