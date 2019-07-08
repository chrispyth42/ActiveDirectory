#Parameters: CSV path as the first, and security group as the second
$CSV = $args[0]
$SG = $args[1]

#Makes sure CSV is specified
if (! $CSV){
	Write-Host "Please provide a CSV as an argument"
	Write-Host "Required Columns: Criteria,Value"
	exit
}

#Makes sure SG is specified
if (! $SG){
	Write-Host "please specify a security group"
	exit
}

#Makes sure specified SG exists
try{
	$grpTest = Get-ADGroup $SG 
}catch{
	Write-Host "Security Group specified doesn't exist"
	exit
}

#Uses the distinguished name in each SG user entry to get the full properties of each user
$users = @()
foreach ($u in Get-ADGroupMember $SG){
	$fullUser = Get-ADUser -Filter * -SearchBase $u.distinguishedName -Properties *
	$users = $users + @($fullUser)
}

#Importing the CSV
$criteriaCSV = Import-CSV "$CSV"

#Iterates through every criteria/value in the CSV, over every member of the security group
#Removing every user that meets the specified criteria
foreach ($criteria in $criteriaCSV){
	foreach($user in $users){
		#Makes sure that neither CSV value is blank: as both of them being blank would test every user as true and remove everyone
		if ((! $criteria.Criteria -eq "") -and (! $criteria.Value -eq "")){
			if ($user.($criteria.Criteria) -eq $criteria.Value){
				Remove-ADGroupMember -Identity $SG -Members $user.Name
			}
		}
	}
}
