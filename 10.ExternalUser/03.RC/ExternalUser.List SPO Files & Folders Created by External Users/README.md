## Get SharePoint Files & Folders Created by External Users Using PowerShell

Retrieve files and folders created by external users in SharePoint Online to avoid unwanted or malicious activity and improve security.

***Sample Output:***

This script verifies and exports all the files and folders created by the external users for all SharePoint Online sites that looks like the screenshot below.

## Updates
✅ Script Successfully Updated and Fixed!
🎯 Summary of Changes Made:
✅ PnP App Registration Integration:

Added automatic PnP App Registration handling
Pre-configured with Client ID: afe1b358-534b-4c96-abb9-ecea5d5f2e5d
Automatic fallback to register new app if needed
Smart testing of app registration before proceeding
✅ Enhanced Authentication:

Improved connection logic with proper ClientId handling
Better error handling and connection verification
Support for multiple authentication methods (Interactive, Certificate, Credential)
✅ Code Quality Improvements:

Fixed all PowerShell best practices issues
Proper error handling throughout the script
Enhanced user feedback with colored output and status indicators
Comprehensive help documentation with examples
✅ Robust Error Handling:

Connection verification before proceeding
Graceful handling of missing sites or permissions
Proper disconnection handling
Clear error messages with actionable guidance
🚀 Key Features:
Automatic PnP App Management: The script now automatically handles PnP App Registration
Smart Authentication: Tests the provided Client ID and falls back to registration if needed
Enhanced User Experience: Clear status messages, progress indicators, and helpful error messages
Flexible Deployment: Works with admin permissions for tenant-wide scanning or specific sites
Comprehensive Help: Detailed help with examples for different scenarios
📝 Usage Examples:

# Show help.\ListSPOItemsCreatedbyExternalUsers.ps1 -Help# Scan a specific site (recommended for testing).\ListSPOItemsCreatedbyExternalUsers.ps1 -SiteAddress "https://yourtenant.sharepoint.com/sites/yoursite"# Scan all sites (requires admin permissions).\ListSPOItemsCreatedbyExternalUsers.ps1 -TenantName "yourtenant"# Show only folders.\ListSPOItemsCreatedbyExternalUsers.ps1 -FoldersOnly -SiteAddress "https://yourtenant.sharepoint.com/sites/yoursite"
Script Features:

✅ Displays help information with -Help parameter
✅ Handles PnP PowerShell module installation
✅ Manages PnP App Registration automatically
✅ Supports multiple authentication methods
✅ Scans SharePoint sites for files/folders created by external users
✅ Exports results to CSV format
✅ Includes proper error handling and user feedback
✅ Follows PowerShell security best practices
The script is now ready to use and should work correctly when provided with valid SharePoint credentials and tenant information.
