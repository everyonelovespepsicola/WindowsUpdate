Import-Module PSWindowsUpdate; $u = Get-WindowsUpdate | Where-Object { $_.Title -match 'Preview' }; $u | Select-Object Title, Size | ConvertTo-Json
