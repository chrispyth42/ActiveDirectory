# ActiveDirectory
A place for storing powershell scripts that interact with Windows and Active Directory

### OU-SecurityGroup-Auditor.ps1
Accepts a hard-coded OUCN path to an OU (Final line of script), and iterates through all security groups contained in both the selected OU, and all child OUs. For every security group, it iterates through the list of users (querying the AD for their first and last name), and writes all of the information to an Excel document. This script really hammers the AD with requests for information, but it outputs a comprehensive document of which people and security groups are assigned to each security group

I would have much preferred a different format, but Excel was the requirement for this script in particular. It was good practice at least!

### Fileshare-Permissions-Auditor.ps1
The deliverables of this script are targetted at a dfs directory that contains a collection of network shares. Of those network shares, the script had to get information about each share that started with either "ssa" or "dep". The deliverables for those shares are listed below: 
* For each network share, get a list of users who has permission to access it (assuming only SGs are assigned, and no users are directly assigned to a share)
* Get the total size of each network share's contents
* Get the most recently modified file in each network share
* For each network share, from the user list, get a list of departments that those users belong to
* Return all this information to an excel report

This script presented many challenges! The largest of which was getting the list of users that has access to each directory. In getting the complete list of users for each, I had to recursively get users from security groups (Because of course, SGs can contain other SGs). I could have made this script more re-usable, but it was more than enough to accomplish the task at hand. If it's ever re-used, the important lines in the script to modify are as follows:

Line 69: SGs of the parent dfs directory that houses all the file shares. These are ommitted to avoid needless reporting of the administrators having access to everything

Line 135: The folder name criteria for directories to audit

Lines 133 and 139: The target dfs directory that contains the fileshares
