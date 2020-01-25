#Given a CN, extract what server the AD object is from (The left-most DC in the FQDN string)
function getServer($UserCN){
    #Split the CN on the string
    $UserCN = $UserCN.split(",")
    #Filter out Non-DCs
    $UserCN = $UserCN | ? {$_ -like 'DC*'}
    #Return the value of the first DC (The server the user is on)
    return $UserCN.split('=')[1]
}

#Recursively gets all users who belong to a security group
function getGroupUsers{
    param($sGroup,$server)
    $output = @()

    #Get members of the input security group. Ignoring builtin and Administrator groups
    try{
        if (($sGroup -ne "Administrators") -and ($server -ne "BUILTIN")){
            $members = (Get-ADGroup $sGroup -Server $server -properties Members).Members
        }else{
            return(@())
        }
    }catch{
        Write-Host ("   Connection failed to $server\$sGroup")
        return(@())
    }
    #Iterate through each, and act differently depending on object type
    foreach ($member in $members){

        #Get object type
        try{
            $objType = (Get-ADObject $member -Server (getServer($member))).objectclass
        }catch{
            Write-Host("   Failed to get object: $member")
            $objType = "None"
        }

        #If user, add to the output list
        if ($objType -eq "user"){
            $output += $member
        }

        #If group, recursively do this function, and append that output list to the main one
        elseif ($objType -eq "group"){
            $output += (getGroupUsers -sGroup $member -server (getServer($member)))
            Write-Host("  $member")
        }
    }
    return $output
}


#Gets a list of user FQDNs who have permissions to a directory
function getDirUsers{
    param($path)
    #Create array to dump users into
    $userList = @()

    #Get all groups that have access to a folder
    $groups = (Get-ACL $path).Access

    #Iterate through those groups
    foreach ($group in $groups){

        #Split the group Identity reference into Server and name
        $groupPath = $group.IdentityReference.Value.split('\')

        #Ignore the two groups that have global access to this whole directory anyways. (If this script is used later, these will need to be changed to be not hardcoded)
        if (($groupPath[1] -ne "ShareManager") -and ($groupPath[1] -ne "DiskAdmins")){

            #Recursively grab user members of the DFS security groups
            Write-Host($groupPath[0] + "\" + $groupPath[1])
            $userList += getGroupUsers -sGroup $groupPath[1] -server $groupPath[0]
        }
    }
    return $userList
}

#Gets size of directory, and its most recent file (Any other desired directory information should be searched for here, so it doesn't do the big Get-ChildItem more than once)
function getSizeAndLastModified{
    param($path)
    #Recursively get contents of path
    $dirContents = Get-ChildItem $path -recurse
    
    #Compute size of directory
    $size = ($dirContents | Measure-Object -Property Length -Sum).Sum/1MB
    $size = [math]::Round($size,2) 

    #Get most recently modified file of directory
    $modified = ($dirContents | sort LastWriteTime | select -last 1).LastWriteTime
    
    return ($size,$modified)
}

#Accepts list of user FQDNs, and lifts the department variable from each; returning a list of Department names
function getDepts{
    param($userList)
    $deptList = @()

    #Iterate through all users in the list, 
    foreach ($user in $userList){
        $server = getServer($user)
        $dept = (Get-ADUser -Identity $user -Server $server -Properties department).department
        $deptList += $dept
    }

    #Remove dupes and return the department list
    return $deptList | select -Unique
}

#Accepts a user FQDN, and returns their full name, and Server\username
function getCleanUserName{
    param($userFQDN)
    $user = Get-ADUser $userFQDN -Server (getServer($userFQDN))
    $fullName = ($user.GivenName + " " + $user.Surname)
    $uName = ((getServer($userFQDN)) + "\" + $user.name)
    return ($fullName,$uName)
}

#Do the thing
function main{
    #Opening an instance of Excel for output
    $excel =  New-Object -ComObject Excel.Application
    $excel.Visible = $true
    $book = $excel.Workbooks.Add()
    $sheet = $book.Sheets.Item(1)

    #Pointers for writing to cells
    $rPtr = 1
    $cPtr = 1

    #Get directory listing, and iterate through each folder that starts with ssa or dep
    $folderList = (Get-ChildItem '\\shares\engineeringShares$').name
    foreach ($folder in $folderList){
        if (($folder -like "ssa*") -or ($folder -like "dep*")){
            Write-Host("-------$folder-------")

            #Assemble folder name
            $folderName = ("\\shares\engineeringShares$\" + $folder)

            #Get that sweet data
            $sizeUpdated = getSizeAndLastModified -path $folderName
            $users = (getDirUsers -path $folderName) | Select -Unique | sort 
            $depts = getDepts -userList $users
            
            #Write filename
            $sheet.cells($rPtr,$cPtr).Value() = $folder
            while ($excel.ready -eq $false){
                Start-Sleep -Milliseconds 10
            }
            $rPtr += 1

            #Write Size
            $sheet.Cells($rPtr,$cPtr).Value() = ("Size: " + $sizeUpdated[0] + "MB")
            while ($excel.ready -eq $false){
                Start-Sleep -Milliseconds 10
            }
            $rPtr += 1

            #Write Updated
            $sheet.Cells($rPtr,$cPtr).Value() = ("Updated: " + $sizeUpdated[1])
            while ($excel.ready -eq $false){
                Start-Sleep -Milliseconds 10
            }
            $rPtr += 2

            #Write Departments Header
            $sheet.Cells($rPtr,$cPtr).Value() = "Departments:"
            while ($excel.ready -eq $false){
                Start-Sleep -Milliseconds 10
            }
            $rPtr += 1

            #Write Departments
            foreach ($dept in $depts){
                $sheet.Cells($rPtr,$cPtr).Value() = $dept
                while ($excel.ready -eq $false){
                    Start-Sleep -Milliseconds 10
                }
                $rPtr += 1
            }
            $rPtr += 1

            #Write Users Header
            $sheet.Cells($rPtr,$cPtr).Value() = "Users:"
            while ($excel.ready -eq $false){
                Start-Sleep -Milliseconds 10
            }
            $rPtr += 1

            #Write Users
            foreach ($user in $users){
                $userClean = getCleanUserName($user)
                $sheet.Cells($rPtr,$cPtr).Value() = $userClean[0]
                while ($excel.ready -eq $false){
                    Start-Sleep -Milliseconds 10
                }
                $sheet.Cells($rPtr,($cPtr + 1)).Value() = $userClean[1]
                while ($excel.ready -eq $false){
                    Start-Sleep -Milliseconds 10
                }
                $rPtr += 1
            }
            #After writing a folder, shift the write pointer accordingly
            $cPtr += 3
            $rPtr = 1 
        }

    }
}

#run the main function
main
