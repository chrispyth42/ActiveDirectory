#Security group parameter
$SG = $args[0]

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

#Iterate through each user, getting full information 
foreach ($u in Get-ADGroupMember $SG){
	$user = Get-ADUser -Filter * -SearchBase $u.distinguishedName -Properties *
	
	#If user distinguishedName isn't within a certain DC, remove them from the SG
	if( ! (($user.distinguishedName -like "*DC=domain*") -and ($user.distinguishedName -like "*DC=local*")) ){
		Write-Host($user.Name)	
		#Remove-ADGroupMember -Identity $SG -Members $user.Name
	}
}
