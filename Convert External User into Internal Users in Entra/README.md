## Bulk Convert External Users to Internal Users in Microsoft 365

The Microsoft Entra admin center and Graph API don't support external user conversion in bulk. This PowerShell script handles bulk conversion of external to internal users by importing CSV while preserving access and existing permissions.

***Sample Output:***  

While converting external users, the script generates a detailed log of each operation to track conversion result, new UPNs, passwords, and other key details for auditing purposes. The output CSV file is password-protected to ensure sensitive information like passwords remains secure. The script exports a CSV file similar to the screenshot below: 

![External to internal user conversion report](https://o365reports.com/wp-content/uploads/2026/03/Sample-Output-1-1024x308.png?v=1774348264)

## Microsoft 365 Management Tool by AdminDroid

Need more than what this script offers? Explore the [AdminDroid Microsoft 365 management tool](https://admindroid.com/?src=GitHub) to access 3500+ reports, 100+ dashboards, and 450+ management actions.

*Simplify conversion of guests to internal users and gain deeper visibility into user management with AdminDroid:*

<https://demo.admindroid.com/#/management/all-actions/6/59>