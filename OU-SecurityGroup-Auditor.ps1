#Fire up new excel to shove the data into
$excel =  New-Object -ComObject Excel.Application
$excel.Visible = $true
$book = $excel.Workbooks.Add()
$sheet = $book.Sheets.Item(1)

#Main pointer for writing to the sheet (Excel indicies start at 1)
$rPtr = 1
$cPtr = 1

#Save the newly initialized excel workbook to a new file
function initSave{
    $date = (Get-Date).ToString().Split(' ')
    $directory = (Get-Location).Path
    $filename = ($date[0] -replace "[/]","-") + "_" + ($date[1] -replace "[:]","") + ".xlsx"
    $book.SaveAs($directory + "\$filename")
    while ($excel.ready -eq $false){
        Start-Sleep -Milliseconds 10
    }
}

#Given a CN, extract what server the AD object is from (The left-most DC in the CN string)
function getServer($UserCN){
    #Split the CN on the string
    $UserCN = $UserCN.split(",")
    #Filter out Non-DCs
    $UserCN = $UserCN | ? {$_ -like 'DC*'}
    #Return the value of the first DC (The server the user is on)
    return $UserCN.split('=')[1]
}

#Provided a security group object, write a block of all its member's details to an excel sheet
function groupWriter($SG){
    #Write the security group name to the sheet
    $sheet.cells($script:rPtr,$script:cPtr).Value() = $SG.Name
    $sheet.cells($script:rPtr,$script:cPtr).Interior.ColorIndex() = 6 
    $script:rPtr += 1
    
    #Write the current SG name to the console to indicate progress
    Write-Host($SG.Name)

    #Get the members of the security group
    $Members = Get-ADGroupMember $SG.DistinguishedName

    #For each member in the group
    foreach($Member in $Members){
        if($Member.objectClass -eq 'user'){
            #Get the server they're in
            $server = getServer($Member.distinguishedName)
            #Use that server to fetch their details (to grab first/last name)
            $userdata = Get-ADUser $Member.distinguishedName -Server $server
            #Write that data to the excel sheet (Sleeping until excel is ready after each write)
            $sheet.Cells($script:rPtr,$script:cPtr).Value() = ($userdata.GivenName + ' ' + $userdata.Surname)
            while ($excel.ready -eq $false){
                Start-Sleep -Milliseconds 10
            }
            $sheet.Cells($script:rPtr,$script:cPtr + 1).Value() = $server + '\' + $Member.name
            while ($excel.ready -eq $false){
                Start-Sleep -Milliseconds 10
            }
            #Increment the row pointer
            $script:rPtr += 1
        }elseif($Member.objectClass -eq 'group'){
            $sheet.Cells($script:rPtr,$script:cPtr).Value() = $Member.Name
            while ($excel.ready -eq $false){
                Start-Sleep -Milliseconds 10
            }
            $sheet.Cells($script:rPtr,$script:cPtr + 1).Value() = 'Security Group'
            while ($excel.ready -eq $false){
                Start-Sleep -Milliseconds 10
            }
            $script:rPtr += 1
        }
    }
    $script:rPtr += 1
}

#Accepts an OU's CN, and outputs its security group structures to an Excel spreadsheet
#The $MainOUName parameter is used by recursion to print the current OU's path
function OUGroups2Excel{
    param([string]$OUCN,[string]$MainOUName)

    #Grab AD Groups inside the OU
    $OUGroups = Get-ADGroup -Filter * -Searchbase $OUCN
    #Grab other OUs inside the OU
    $NestedOUs = Get-ADOrganizationalUnit -Filter * -SearchBase $OUCN
    $NestedOUs = $NestedOUs | Where-Object {$_.DistinguishedName -ne $OUCN} #Remove the current OU from the results so it doesn't recurse forever
    #Grab the current OUs details
    $CurrentOU = Get-ADOrganizationalUnit $OUCN

    #Write a header for the block containing the OU's name (With an if to handle if the $MainOUName argument exists yet)
    if($MainOUName){
        $sheet.cells($script:rPtr,$script:cPtr).Value() = $MainOUName
    }else{
        $sheet.cells($script:rPtr,$script:cPtr).Value() = $CurrentOU.Name
        $MainOUName = $CurrentOU.Name
    }
    #Make the cell green
    $sheet.cells($script:rPtr,$script:cPtr).Interior.ColorIndex() = 4
    #Increment the row pointer
    $script:rPtr += 1

    #Write each of the current OUs security groups and their members to the sheet
    foreach($Group in $OUGroups){
        groupWriter($Group)
    }
    
    #Save the book after each OU
    $book.Save()
    while ($excel.ready -eq $false){
        Start-Sleep -Milliseconds 10
    }
    
    #Use the power of   R E C U R S I O N   to grab all sub-OUs as well
    foreach($OU in $NestedOUs){
        #Shift the pointer 3 columns right, and reset the row pointer to 1 before repeating with the nested OU
        $script:rPtr = 1
        $script:cPtr += 3
        OUGroups2Excel -OUCN $OU.DistinguishedName -MainOUName ($MainOUName + "\" + $OU.Name)
    }
}

#Saves the newly generated workbook into the working directory
initSave

#Runs the main function
OUGroups2Excel -OUCN 'OU=itadmin,DC=itdepartment,DC=ad,DC=samplesite,DC=com'
