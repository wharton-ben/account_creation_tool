<#	
	.NOTES
	===========================================================================
	 Created with: 	SAPIEN Technologies, Inc., PowerShell Studio 2021 v5.8.187
	 Created on:   	6/18/2021 3:42 PM
	 Created by:   	Ben
	 Organization: 	
	 Filename:     	
	===========================================================================
	.DESCRIPTION
		A description of the file.
#>

#region - Set-Password Function
Function Set-Password
{
	<#The $Dictionary and $Dictionary2 variables will contain a random word from a list that is stored in a .txt file
	*The $Special variable will contain a special character from a list
	*The $Number variable will contain a random number from 0 - 9
	*If the password is too short, a number will be appended to the password until the password meets the requirements from the domain
	*Once the password is the correct length, a special character will be added to the end of the password.
	#>
	$pwlength = $config.'pwlength'
	$Password = $null
	
	#Grab the words from the dictionary file and create a password
	$Dictionary = Get-Content -Path $config.'dictionarypath' | Get-Random -Count 1
	$Dictionary2 = Get-Content -Path $config.'dictionarypath' | Get-Random -Count 1
	$Special = Get-Random ('!', '?')
	$Number = Get-Random -Minimum 0 -Maximum 10
	$Password = $Dictionary + $Dictionary2 + $Number
	if ($Password.length -le ($pwlength - 1))
	{
		do
		{
			$Number = Get-Random -Minimum 0 -Maximum 10
			$Password = $Password + $Number
		}
		while ($Password.length -lt ($pwlength - 1))
	}$Password = $Password + $Special
	return $Password
}
#endregion

#region - Add-VAAccount
function Add-VAAccount
{
	param
	(
		[parameter(Mandatory = $true)]
		[string]$filename
	)
	import-module activedirectory
	
	#Description: This script pulls from a CSV file located on the user's local hard drive
	
	#Region - Initial import of the CSV, declares an area to hold account details in
	$csv = Import-Csv -Path $filename
	$config = Import-Csv -path .\config.csv
	$groups = $groupselection.CheckedItems
	$results = @()
	$pwlength = $($config.pwlength)
	#endregion
	
	#Region - The Foreach loop which each user's details will be run through
	ForEach ($user in $csv)
	{
		#$Path = "OU=$($comboboxSelectDomain.Text),OU=Tier3,OU=USVA-DC,OU=Accounts,OU=Clients,DC=LCAHNCRKC,DC=net"
		$Path = "OU=$($comboboxSelectDomain.text)$($config.path)"
		$results = @()
		$pwlength = $($config.pwlength)
		
		#region - Declares the variables for use in the script
		$UserFirstname = $user.FirstName
		$UserLastname = $user.LastName
		$Initials = $user.MiddleName
		$join = $user.FirstName + $user.LastName
		$UserDisplayName = "$Userlastname, $UserFirstName $Initials"
		$userID = $user.Username
		$userPrincipal = $userID + $($config.upn)
		$Description = $user.Credential
		$Email = $user.EmailAddress
		$StreetAddress = $user.StreetAddress
		$City = $user.City
		$State = $user.State
		$PostalCode = $user.PostalCode
		$OfficePhone = $user.'Phone Number'
		$accountExpiration = (get-date).AddDays(365)
		$ADpwd = (Set-Password)
		#endregion
		
		#region - if user exists, skip. If not, create the user.
		
		#This if statement checks to see if the user exists
		$duplicate = Get-ADUser -LDAPFilter "(SamAccountName=$userid)"
		If ($duplicate -eq $null)
		{
			#Creates the new AD User
			New-ADUser -Name $UserDisplayName `
					   -Path $Path `
					   -GivenName $UserFirstname `
					   -Surname $UserLastname `
					   -Initials $Initials `
					   -SamAccountName $userID `
					   -AccountPassword (ConvertTo-SecureString $ADpwd -AsPlainText -Force) `
					   -ChangePasswordAtLogon $true `
					   -Enabled $true `
					   -Description $Description `
					   -UserPrincipalName $userPrincipal `
					   -DisplayName $UserDisplayName `
					   -EmailAddress $Email `
					   -StreetAddress $StreetAddress `
					   -City $City `
					   -State $State `
					   -PostalCode $PostalCode `
					   -OfficePhone $OfficePhone `
					   -AccountExpirationDate $accountExpiration `
					   -CannotChangePassword $false `
					   -SmartcardLogonRequired $false `
			
			
			#This is where the user settings such as 'password never expires' or 'password must be changed upon first log on' will be specified
			#Set-ADUser -Identity $userID 
			
			#Add user to groups - Groups will be determined by which groups are listed in the groups.csv file.
			foreach ($group in $groups)
			{
				Add-ADGroupMember $group -Members $userID
			}
			$Action = "$userID created sucessfully."
		}
		Else
		{
			#This will grab the top level OU from the existing account and display the OU in the results file. 
			$existingOU = Get-ADUser $userID -Properties * | Select-Object distinguishedname
			$existingOU = $existingOU.distinguishedname
			$existingOU = $existingOU -split ('OU=')
			$existingOU = [regex]::Replace($existingOU[1], "[^a-zA-Z^0-9\s=]", "")
			$Path = $existingOU
			$ADpwd = "User already exists. Password was not reset."
			$Action = "Account $UserID already exists"
		}
		#endregion
		
		#region - Details stored in the $details variable will be exported to the VAACTResults.csv file. 
		$details = [ordered]@{
			"OU"		   = $Path;
			"First"	       = $UserFirstname;
			"Last"		   = $UserLastname;
			"Username"	   = $userID;
			"Email"	       = $Email;
			"EDIPI"	       = $EDIPI;
			"Password"	   = $ADpwd;
			"Action Taken" = $Action;
		}
		
		#Add the date to the file
		$resultdate = Get-Date -Format "hh-mm-MM-dd-yyyy"
		$resultfile = "$($config.result_file)\VAACTresults$($resultdate).csv"
		
		#Add the details for each user to the results file
		$results += New-Object System.Management.Automation.PSObject -property $details | Export-Csv $resultfile -NoTypeInformation -force -Append
		#endregion
	}
	#endregion
	
	#region - script completed, launch results file.
	
	#The user will receive notification that the users have been created. The message on the screen will prompt them to check for details in the results file. Results will launch automatically. 
	$note.popup("Complete, check the contents of the VAACT results file for details.")
	
	#Launch the results file
	Start-Process $resultfile
	#endregion
}
#endregion



