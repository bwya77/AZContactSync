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
    Write-Verbose "Connecting to Graph API"
      $ReqTokenBody = @{
          Grant_Type    = "client_credentials"
          Scope		  = "https://graph.microsoft.com/.default"
          client_Id	  = $clientID
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
    Write-Verbose "Creating new contact folder: $Name"
  $body = @"
{
    "displayName": "$Name"
}
"@
  }
  Process {
    $request = @{
      Method = "Post"
      Uri    = "https://graph.microsoft.com/v1.0/users/$userprincipalName/contactFolders"
      ContentType = "application/json"
      Headers = @{ Authorization = "Bearer $($AccessToken)" }
      Body   = $Body
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
Write-Verbose "Getting contact folders for $UserPrincipalName"
}
Process {
  $request = @{
    Method = "Get"
    Uri    = "https://graph.microsoft.com/v1.0/users/$UserPrincipalName/contactFolders/"
    ContentType = "application/json"
    Headers = @{ Authorization = "Bearer $($AccessToken)" }
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
    Write-Verbose "Getting list items from $listID"
      $headers = @{
          Authorization = "Bearer $accessToken"
      }
      $apiUrl = "https://graph.microsoft.com/v1.0/sites/$siteID/lists/$listID/items?expand=fields"
  }
  process {
      $listItems = Invoke-RestMethod -Uri $apiURL -Headers $headers -Method GET
  }
  end {
      return $listItems.value.fields
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
      Write-Verbose "Creating new contact in contact folder: $contactFolderID"
      $request = @{
        Method = "Post"
        Uri    = "https://graph.microsoft.com/v1.0/users/$UserPrincipalName/contactFolders/$contactfolderID/contacts"
        ContentType = "application/json"
        Headers = @{ Authorization = "Bearer $($accessToken)" }
        Body   = $Body
      }
    }
Else {
  Write-Verbose "Creating new contact outside of contact folder"
  $request = @{
    Method = "Post"
    Uri    = "https://graph.microsoft.com/v1.0/users/$UserPrincipalName/contacts"
    ContentType = "application/json"
    Headers = @{ Authorization = "Bearer $($accessToken)" }
    Body   = $Body
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
      Method = "PATCH"
      Uri    = "https://graph.microsoft.com/v1.0/users/$UserPrincipalName/contacts/$matchcontactID"
      ContentType = "application/json"
      Headers = @{ Authorization = "Bearer $accessToken" }
      Body = $body
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
}
Process {
  if ($contactfolderID)
  {
    $request = @{
      Method = "Get"
      Uri    = "https://graph.microsoft.com/v1.0/users/$UserPrincipalName/contactFolders/$contactfolderID/contacts"
      ContentType = "application/json"
      Headers = @{ Authorization = "Bearer $($AccessToken)" }
    }
  }
  Else {
    $request = @{
      Method = "Get"
      Uri    = "https://graph.microsoft.com/v1.0/users/$UserPrincipalName/contacts"
      ContentType = "application/json"
      Headers = @{ Authorization = "Bearer $($AccessToken)" }
    }

  }
}
End {
  Invoke-RestMethod @request
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
  }
  Process {
    If ($contactFolderID) {
      Write-Verbose "Removing contact in contact folder: $contactFolderID"
      $request = @{
        Method = "DELETE"
        Uri    = "https://graph.microsoft.com/v1.0/users/$UserPrincipalName/contactFolders/$contactfolderID/contacts/$contactID"
        ContentType = "application/json"
        Headers = @{ Authorization = "Bearer $($accessToken)" }
      }
    }
Else {
  Write-Verbose "Removing contact outside of contact folder"
  $request = @{
    Method = "DELETE"
    Uri    = "https://graph.microsoft.com/v1.0/users/$UserPrincipalName/contacts/$contactID"
    ContentType = "application/json"
    Headers = @{ Authorization = "Bearer $($accessToken)" }
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
  [string]$AccessToken
  )
  Begin { 
    Write-Verbose "Getting all users"
  }
  Process {
    $request = @{
      Method = "Get"
      Uri    = "https://graph.microsoft.com/v1.0/users/"
      ContentType = "application/json"
      Headers = @{ Authorization = "Bearer $($accessToken)" }
    }
  }
  End {
    Invoke-RestMethod @request
  }
}

[system.string]$contactfolderName = 'Work Contacts'
$VerbosePreference = 'Continue'

$clientId = Get-AutomationVariable -Name "clientID"
$tenantID = Get-AutomationVariable -Name "tenantID"
$clientSecret = Get-AutomationVariable -Name "clientSecret"
$siteID = Get-AutomationVariable -Name "siteID"
$listID = Get-AutomationVariable -Name "listID"


$token = Connect-GraphAPI -clientID $clientID -tenantID $tenantID -clientSecret $clientSecret

#Get all users that have a mail attribute
$users = (Get-Users -accessToken $token.access_token).value | Where-Object {$null -ne $_.mail}
foreach ($user in $users) {
  $userprincipalName = $user.userprincipalName


  $ContactFolders = Get-ContactFolders -UserPrincipalName $userprincipalName -AccessToken $token.access_token
  if ($ContactFolders.displayName -notcontains $contactfolderName) {
    $workContactsID = (New-ContactFolder -Name $contactfolderName -UserPrincipalName $userprincipalName -AccessToken $token.access_token).id
  }
  else {
    $workcontactsID = ($ContactFolders | Where-Object {$_.displayName -eq "$contactfolderName"}).id
  }
  Write-Verbose "Work Contacts ID: $workContactsID"
  #Get list items and iterate through them
  $listItems = Get-ListItems -siteID $siteID -listID $listID -accessToken $Token.access_token
  foreach ($Item in $listItems) {
    #Check if the contact exists in the user's contacts
    Write-Verbose "Getting all user contacts"
    $userContacts = Get-Contacts -UserPrincipalName $userprincipalName -contactFolderID $workContactsID -AccessToken $token.access_token 

      $Match = ($userContacts).value | Where-Object { $_.businessPhones -contains $item.phoneNumber } | Select-Object -First 1
      #If the contact phone number is present in the users contacts already, check if the first and last names match
      If ($Match.givenName -eq $item.title -and $Match.surname -eq $item.surname) {
        #If the first name and last name match, the contact does not need further updating
        Write-Verbose "first and last names match for contact: $($item.phoneNumber)"
      }
      #If either the firstname or lastname don't match, update the contact
      Elseif ($Match.givenName -ne $item.title -or $Match.surname -ne $item.surname -and $Null -ne $Match) {
        Write-Verbose "The firstname or the lastname for the contact $($item.phoneNumber) do not match. Updating contact"
        Set-Contact -UserPrincipalName $userprincipalName -accessToken $token.access_token -givenName $Item.title -surname $Item.surname -businessPhone $Item.phoneNumber -matchcontactID $Match.id
      }
      #If there is no matching contact, we must create a new contact 
      Else {
        Write-Verbose "No matching contact found for $($item.phoneNumber). Creating new contact"
        New-Contact -UserPrincipalName $userprincipalName -AccessToken $token.access_token -givenName $Item.title -surname $Item.surname -businessPhone $Item.phoneNumber -contactFolderID $workContactsID
      }
  }

  #Refresh the list of contacts and list items so we are working with the most current data 
  $userContacts = (Get-Contacts -UserPrincipalName $userprincipalName -contactFolderID $workContactsID -AccessToken $token.access_token).value
  $listItems = Get-ListItems -siteID $siteID -listID $listID -accessToken $Token.access_token

  #Get all contacts that are not in the SharePoint list
  $removeContacts = $userContacts | Where-Object { ($_.givenName -notin $listItems.title ) -or ($_.surname -notin $listItems.surname)}

  foreach ($item in $removeContacts) {
    Write-Verbose "Removing contact: givenName: $($item.givenName) surname: $($item.surname)"
    Remove-Contact -UserPrincipalName $userprincipalName -accessToken $token.access_token -contactID $item.id -contactFolderID $workContactsID
  }
}
