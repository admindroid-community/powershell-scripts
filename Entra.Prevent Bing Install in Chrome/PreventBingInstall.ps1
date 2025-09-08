$RegistryPath = "HKLM:\SOFTWARE\Policies\Microsoft\Office\16.0\Common\Officeupdate"
$Name = "preventbinginstall"
$value = "00000001"

if(!(Test-Path $RegistryPath))
{
 Write-Host Office 365 Pro Plus not available -ForegroundColor Yellow
}
else
{
 New-ItemProperty -Path $registryPath -Name $name -Value $value -PropertyType DWORD -Force | Out-Null
 Write-Host Successfully added registry key to prevent Bing install in Chrome -ForegroundColor Green
}
