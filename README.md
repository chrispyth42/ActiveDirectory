# ActiveDirectory
A place for storing powershell scripts that interact with Windows Active Directory

### OU-SecurityGroup-Auditor.ps1
Accepts a hard-coded OUCN path to an OU, and iterates through all security groups contained in both the selected OU, and all child OUs. For every security group, it iterates through the list of users (querying the AD for their first and last name), and writes all of the information to an Excel document. This script really hammers the AD with requests for information, but it outputs a comprehensive document of which people and security groups are assigned to each security group

I would have much preferred a different format, but Excel was the requirement for this script in particular. It was good practice at least!
