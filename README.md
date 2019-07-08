# ActiveDirectory
A place for storing powershell scripts that interact with Windows Active Directory

### RemoveSGUsers.ps1
Is called in the format .\RemoveUsers.ps1 CSV_PATH 'SG NAME' . It iterates through every user listed in a security group, and gets their full information using Get-ADUser. After doing so, it goes through each criteria in the CSV (see data\criteria.csv), and removes every user from the specified security group that meets the properties listed there. I'm fairly sure it's reliable, but it has NOT BEEN TESTED very extensively, so feel free to test and judge it before using it. In the sample CSV I've provided, it removes users named 'ike', as well as users that were created at that specific time, from the Security Group

### RemoveSGUsers-Simple.ps1
Is like the non-simple version, except instead of reading in removal criteria from a CSV, it's all within the script itself. This script just iterates through the members of a Security Group, and prints/removes them if whatever specified condition is met
