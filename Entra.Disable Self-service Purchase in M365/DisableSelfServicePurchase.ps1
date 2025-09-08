<#
=============================================================================================
Name:              Disable self-service purchase for Microsoft 365 products
Description:       This script disables self-service purchase/ trial capability for all the supported M365 products
Script written by: AdminDroid community team
Website:           admindroid.com
=====================================================================================
#>

Install-Module -Name MSCommerce -Scope CurrentUser #once you install you should remove this line
Import-Module -Name MSCommerce 
Connect-MSCommerce #sign-in with your global or billing administrator account when prompted
Get-MSCommerceProductPolicies -PolicyId AllowSelfServicePurchase | Where { $_.PolicyValue -eq “Enabled”} | forEach { 
 Update-MSCommerceProductPolicy -PolicyId AllowSelfServicePurchase -ProductId $_.ProductID -Enabled $false  
}