Function Connect-GraphAPI {
  [CmdletBinding()]
  Param (
    [Parameter(Mandatory)]
    [string]$clientID,
    [Parameter(Mandatory)]
    [string]$tenantID,
    [Parameter(Mandatory)]
    [string]$clientSecret
  )
  begin {
    Write-Output "Connecting to Graph API"
    $ReqTokenBody = @{
      Grant_Type    = "client_credentials"
      Scope         = "https://graph.microsoft.com/.default"
      client_Id     = $clientID
      Client_Secret = $clientSecret
    }
  }
  process {
    $tokenResponse = Invoke-RestMethod -Uri "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token" -Method POST -Body $ReqTokenBody
  }
  end {
    return $tokenResponse
  }
}
Function New-ContactFolder {
  [CmdletBinding()]
  Param (
    [Parameter(Mandatory)]
    [string]$Name,
    [Parameter(Mandatory)]
    [string]$AccessToken,
    [Parameter(Mandatory)]
    [string]$UserPrincipalName
  )
  Begin {
    Write-Output "Creating new contact folder: $Name"
    $body = @"
  {
      "displayName": "$Name"
  }
"@
  }
  Process {
    $request = @{
      Method      = "Post"
      Uri         = "https://graph.microsoft.com/v1.0/users/$userprincipalName/contactFolders"
      ContentType = "application/json"
      Headers     = @{ Authorization = "Bearer $($AccessToken)" }
      Body        = $Body
    }
    $Post = Invoke-RestMethod @request
  }
  End {
    return $Post
  }
}
Function Get-ContactFolders {
  Param (
    [Parameter(Mandatory)]
    [string]$UserPrincipalName,
    [Parameter(Mandatory)]
    [string]$AccessToken
  )
  Begin { 
    Write-Output "Getting contact folders for $UserPrincipalName"
  }
  Process {
    $request = @{
      Method      = "Get"
      Uri         = "https://graph.microsoft.com/v1.0/users/$UserPrincipalName/contactFolders/"
      ContentType = "application/json"
      Headers     = @{ Authorization = "Bearer $($AccessToken)" }
    }
    $Data = Invoke-RestMethod @request
  
  }
  End {
    return $Data.Value
  }
}
Function Get-ListItems {
  [CmdletBinding()]
  Param (
    [Parameter(Mandatory)]
    [string]$siteID,
    [Parameter(Mandatory)]
    [string]$listID,
    [Parameter(Mandatory)]
    [string]$accessToken
  )
  begin {
    $allListItems = @()
    Write-Output "Getting list items from $listID"
    $headers = @{
      Authorization = "Bearer $accessToken"
    }
    $apiUrl = "https://graph.microsoft.com/v1.0/sites/$siteID/lists/$listID/items?expand=fields"
  }
  process {
    $listItems = Invoke-RestMethod -Uri $apiURL -Headers $headers -Method GET
    $allListItems += $listItems.value.fields
    if ($listItems.'@odata.nextLink') {
      do {
        $listItems = Invoke-RestMethod -Uri $listItems.'@odata.nextLink' -Headers @{ Authorization = "Bearer $($AccessToken)" } -Method "Get" -ContentType "application/json"
        $allListItems += $listItems.value.fields
      } Until (!$listItems.'@odata.nextLink')
    }
  }
  end {
    return $allListItems
  }
}
Function New-Contact {
  Param (
    [Parameter(Mandatory)]
    [string]$UserPrincipalName,
    [Parameter(Mandatory)]
    [string]$AccessToken,
    [Parameter(Mandatory)]
    [string]$givenName,
    [Parameter(Mandatory)]
    [string]$surname,
    [Parameter(Mandatory)]
    [string]$businessPhone,
    [Parameter()]
    [string]$contactFolderID
  )
  Begin {
    $body = @"
  {
      "givenName": "$givenName",
      "surname": "$surname",
      "businessPhones": [
        "$businessPhone"
      ]
  }
"@ 
  }
  Process {
    If ($contactFolderID) {
      Write-Output "Creating new contact in contact folder: $contactFolderID"
      $request = @{
        Method      = "Post"
        Uri         = "https://graph.microsoft.com/v1.0/users/$UserPrincipalName/contactFolders/$contactfolderID/contacts"
        ContentType = "application/json"
        Headers     = @{ Authorization = "Bearer $($accessToken)" }
        Body        = $Body
      }
    }
    Else {
      Write-Output "Creating new contact outside of contact folder"
      $request = @{
        Method      = "Post"
        Uri         = "https://graph.microsoft.com/v1.0/users/$UserPrincipalName/contacts"
        ContentType = "application/json"
        Headers     = @{ Authorization = "Bearer $($accessToken)" }
        Body        = $Body
      }
    }
  }
  End {
    Invoke-RestMethod @request
  }
}
Function Set-Contact { 
  Param (
    [Parameter(Mandatory)]
    [string]$UserPrincipalName,
    [Parameter(Mandatory)]
    [string]$accessToken,
    [Parameter(Mandatory)]
    [string]$givenName,
    [Parameter(Mandatory)]
    [string]$surname,
    [Parameter(Mandatory)]
    [string]$businessPhone,
    [Parameter(Mandatory)]
    [string]$matchcontactID
  )
  Begin {
    $body = @"
  {
      "givenName": "$givenName",
      "surname": "$surname",
      "businessPhones": [
        "$businessPhone"
      ]
  }
"@ 
  }
  Process {
    $request = @{
      Method      = "PATCH"
      Uri         = "https://graph.microsoft.com/v1.0/users/$UserPrincipalName/contacts/$matchcontactID"
      ContentType = "application/json"
      Headers     = @{ Authorization = "Bearer $accessToken" }
      Body        = $body
    }
  }
  End {
    Invoke-RestMethod @request
  }
  
}
Function Get-Contacts {
  Param (
    [Parameter(Mandatory)]
    [string]$UserPrincipalName,
    [Parameter(Mandatory)]
    [string]$AccessToken,
    [Parameter()]
    [string]$contactFolderID
  )
  Begin { 
    [system.array]$allContacts = @()
  }
  Process {
    if ($contactfolderID) {
      $request = @{
        Method      = "Get"
        Uri         = "https://graph.microsoft.com/v1.0/users/$UserPrincipalName/contactFolders/$contactfolderID/contacts"
        ContentType = "application/json"
        Headers     = @{ Authorization = "Bearer $($AccessToken)" }
      }
    }
    Else {
      $request = @{
        Method      = "Get"
        Uri         = "https://graph.microsoft.com/v1.0/users/$UserPrincipalName/contacts"
        ContentType = "application/json"
        Headers     = @{ Authorization = "Bearer $($AccessToken)" }
      }
    }
    $contacts = Invoke-RestMethod @request
    $allContacts += $contacts.value
    if ($contacts.'@odata.nextLink') {
      do {
        $contacts = Invoke-RestMethod -Uri $contacts.'@odata.nextLink' -Headers @{ Authorization = "Bearer $($AccessToken)" } -Method "Get" -ContentType "application/json"
        $allContacts += $contacts.value
      } Until (!$contacts.'@odata.nextLink')
    }
  }
  End {
    $allContacts
  }
}
Function Remove-Contact {
  Param (
    [Parameter(Mandatory)]
    [string]$UserPrincipalName,
    [Parameter(Mandatory)]
    [string]$AccessToken,
    [Parameter(Mandatory)]
    [string]$contactID,
    [Parameter()]
    [string]$contactFolderID
  )
  Begin {
    Write-Output "Removing contact: $contactID for user $UserPrincipalName"
  }
  Process {
    If ($contactFolderID) {
      Write-Output "Removing contact in contact folder: $contactFolderID"
      $request = @{
        Method      = "Delete"
        Uri         = "https://graph.microsoft.com/v1.0/users/$UserPrincipalName/contactFolders/$contactfolderID/contacts/$contactID"
        ContentType = "application/json"
        Headers     = @{ Authorization = "Bearer $($accessToken)" }
      }
    }
    Else {
      Write-Output "Removing contact outside of contact folder"
      $request = @{
        Method      = "Delete"
        Uri         = "https://graph.microsoft.com/v1.0/users/$UserPrincipalName/contacts/$contactID"
        ContentType = "application/json"
        Headers     = @{ Authorization = "Bearer $($accessToken)" }
      }
    }
  }
  End {
    Invoke-RestMethod @request
  }
}
Function Get-Users {
  Param (
    [Parameter(Mandatory)]
    [system.string]$AccessToken
  )
  Begin { 
    [system.array]$AlluserItems = @()
    Write-Output "Getting all users"
    $APIendpoint = 'https://graph.microsoft.com/v1.0/users?$select=id,displayName,assignedLicenses,assignedPlans,userprincipalname,mail'
  }
  Process {
    $request = @{
      Method      = "Get"
      Uri         = $APIendpoint
      ContentType = "application/json"
      Headers     = @{ Authorization = "Bearer $($accessToken)" }
    }
    $Users = Invoke-RestMethod @request
    $AlluserItems += $Users.value
    if ($Users.'@odata.nextLink') {
      do {
        $Users = Invoke-RestMethod -Uri $listItems.'@odata.nextLink' -Headers @{ Authorization = "Bearer $($AccessToken)" } -Method "Get" -ContentType "application/json"
        $AlluserItems += $AlluserItems += $Users.value
      } Until (!$Users.'@odata.nextLink')
    }
  }
  End {
    $AlluserItems
  }
}

[system.string]$contactfolderName = 'Work Contacts'

$clientId = Get-AutomationVariable -Name "clientID"
$tenantID = Get-AutomationVariable -Name "tenantID"
$clientSecret = Get-AutomationVariable -Name "clientSecret"
$siteID = Get-AutomationVariable -Name "siteID"
$listID = Get-AutomationVariable -Name "listID"



$token = Connect-GraphAPI -clientID $clientID -tenantID $tenantID -clientSecret $clientSecret

[system.int32]$countUsers = 0
#Get all users that have a mail attribute
$users = Get-Users -accessToken $token.access_token | Where-Object {($null -ne $_.mail) -and ($_.assignedLicenses -ne $null)}
foreach ($user in $users) {
  $countUsers ++
  [system.int32]$listcount = 1
  Write-Output "---- Working on user $countUsers of $($users.count) ----"
  $userprincipalName = $user.userprincipalName
  Write-Output "Working on user: $userprincipalName"

  $ContactFolders = Get-ContactFolders -UserPrincipalName $userprincipalName -AccessToken $token.access_token
  if ($ContactFolders.displayName -notcontains $contactfolderName) {
    $workContactsID = (New-ContactFolder -Name $contactfolderName -UserPrincipalName $userprincipalName -AccessToken $token.access_token).id
  }
  else {
    $workcontactsID = ($ContactFolders | Where-Object { $_.displayName -eq "$contactfolderName" }).id
  }
  Write-Output "Work Contacts ID: $workContactsID"
  #Get list items and iterate through them
  $listItems = Get-ListItems -siteID $siteID -listID $listID -accessToken $Token.access_token
  foreach ($Item in $listItems) {
    Write-Output "---- Working on list item $listcount of $($listitems.count) ----"
    #Check if the contact exists in the user's contacts
    Write-Output "Working on $($item.phoneNumber) from SharePoint list"
    $userContacts = Get-Contacts -UserPrincipalName $userprincipalName -contactFolderID $workContactsID -AccessToken $token.access_token 

    Write-output "Checking if contact: $($item.phoneNumber) exists in user contacts"
    $Match = $userContacts | Where-Object { $_.businessPhones -contains $item.phoneNumber } | Select-Object -First 1
    #If the contact phone number is present in the users contacts already, check if the first and last names match
    If ($Match.givenName -eq $item.title -and $Match.surname -eq $item.surname) {
      #If the first name and last name match, the contact does not need further updating
      Write-Output "first and last names match for contact: $($item.phoneNumber)"
    }
    #If either the firstname or lastname don't match, update the contact
    Elseif ($Match.givenName -ne $item.title -or $Match.surname -ne $item.surname -and $Null -ne $Match) {
      Write-Output "The firstname or the lastname for the contact $($item.phoneNumber) do not match. Updating contact"
      Set-Contact -UserPrincipalName $userprincipalName -accessToken $token.access_token -givenName $Item.title -surname $Item.surname -businessPhone $Item.phoneNumber -matchcontactID $Match.id
    }
    #If there is no matching contact, we must create a new contact 
    Else {
      Write-Output "No matching contact found for $($item.phoneNumber). Creating new contact"
      New-Contact -UserPrincipalName $userprincipalName -AccessToken $token.access_token -givenName $Item.title -surname $Item.surname -businessPhone $Item.phoneNumber -contactFolderID $workContactsID
    }
    $listcount++
  }

  #Refresh the list of contacts and list items so we are working with the most current data 
  $userContacts = Get-Contacts -UserPrincipalName $userprincipalName -contactFolderID $workContactsID -AccessToken $token.access_token
  $listItems = Get-ListItems -siteID $siteID -listID $listID -accessToken $Token.access_token

  #Get all contacts that are not in the SharePoint list
  $removeContacts = $userContacts | Where-Object { ($_.givenName -notin $listItems.title ) -or ($_.surname -notin $listItems.surname) }

  foreach ($i in $removeContacts) {
    Write-Output "Removing contact: givenName: $($item.givenName) surname: $($item.surname)"
    Remove-Contact -UserPrincipalName $userprincipalName -accessToken $token.access_token -contactID $i.id -contactFolderID $workContactsID
  }
}
