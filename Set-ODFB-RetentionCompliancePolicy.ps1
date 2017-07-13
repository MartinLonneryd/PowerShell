# Secuirty and Compliance portal Retention Policy for ODFB
$ErrorActionPreference = 'stop'

# Write which host is running script
$env:COMPUTERNAME

# Enables and disables the production mode for the script, if set to $false or no value the script will only log output and not do any changes.
$ProductionMode = $true

# Loading functions
Write-Output 'Loading functions...'

# Function to split and make sense of SPO ImmutableIdentity parameter: 1;Site;https://hennesandmauritz-my.sharepoint.com/personal/firstname_lastname_domain_top;12bf47a2-9396-4d82-9ca5-e2b40239d123
function Select-PolicyUser
{
  <#
      .SYNOPSIS
      Takes the ImmutableIdentity object from the retention policy cmdlet and create a custom object that contains the information that has been gathered
      Describe purpose of "Select-PolicyUser" in 1-2 sentences.

      .PARAMETER InputObject
      The output from (Get-RetentionCompliancePolicy).OneDriveLocation piped to this variable

      .EXAMPLE
      $odfbpolicy.OneDriveLocation | Select-PolicyUser
  #>
  param
  (
    [Parameter(Mandatory=$true, ValueFromPipeline=$true, HelpMessage='Data to process')]
    $InputObject
  )
  process
  {
    
    [pscustomobject]@{
      Namn = $InputObject.ImmutableIdentity.Split(';')[3]
      ObjectID = $InputObject.ImmutableIdentity.Split(';')[-3]
      Site = $InputObject.ImmutableIdentity.Split(';')[2]
    }
    
  }
}

# Function get-msoluser and create objects of each selected value to streamline all user objects for processing
function Select-MSOLUser
{
  <#
      .SYNOPSIS
      Takes the MSOLUser object from the get-msoluser cmdlet and create a custom object that UserPrincipalName,DisplayName,ImutableID,ProxyAddresses and ProxyAddresses without the SMTP: in the begining
      Describe purpose of "Select-PolicyUser" in 1-2 sentences.

      .PARAMETER InputObject
      The MSOLuser object from get-msoluser

      .EXAMPLE
      Get-MsolUser | Select-MSOLUser
  #>


  param
  (
    [Parameter(Mandatory=$true, ValueFromPipeline=$true, HelpMessage='Data to process')]
    $InputObject
  )
  process
  {
    
    [pscustomobject]@{
      UserPrincipalName = $InputObject.UserPrincipalName
      DisplayName = $InputObject.DisplayName
      ImmutableId = $InputObject.ImmutableId
      IsLicensed = $InputObject.IsLicensed
      ProxyAddresses = $InputObject.ProxyAddresses
      ProxyAddressesWithoutSMTP = foreach ($ProxyAddress in $InputObject.ProxyAddresses){
        $ProxyAddress.Substring(5)
      }
    }
    
  }
}

# Function to split and make sense of Get-SPOSite Url and guess owner and match to upn
function Select-SharePointSite
{
  <#
      .SYNOPSIS
      Is used to gather information on a ODFB site, taking in the information of the URL and assuming a owner for the ODFB site
      Checking if the ODFB site end with a number and also checks if the Owner and the URL for the ODFB site matches

      .EXAMPLE
      $SPOSites = Get-SPOSite -Template "SPSPERS#10" -IncludePersonalSite:$true -Limit All
      Write-Output 'Process sites with function Select-SharePointSite'
      $SPOSites = $SPOSites | Select-SharePointSite
  #>


  param
  (
    [Parameter(Mandatory=$true, ValueFromPipeline=$true, HelpMessage='Data to process')]
    $InputObject
  )
  
  process
  {
    $guessUPN = $null
    $EndURLNumber = $InputObject.url.Split('_')[-1]
    $UpnFromUrl = $InputObject.Url.Split('/')[-1]
    $DomainFromUrl = $UpnFromUrl.Split('_')[-2]
    foreach ($obj in $InputObject.Url.Split('/')[-1].Split('_')){
      if ($obj -eq $EndURLNumber){
        $guessUPN += $obj
      }
      elseif ($obj -eq $DomainFromUrl){
        $obj = '@'+"$DomainFromUrl"
        $guessUPN += $obj+'.'
      }
      else{
        $guessUPN += $obj+'.'
      }
      $guessUPN = $guessUPN.Replace('.@','@')
    }   
    [pscustomobject]@{
      Url = $InputObject.Url
      Owner = $InputObject.Owner
      Title = $InputObject.Title
      Template = $InputObject.Template
      OwnerAndUpnMatch = if ($InputObject.Owner -eq $guessUPN){
        $true
      }
      else{
        $false
      }
      SiteWithNumber = try{  
        [int]($InputObject.url.Split('_')[-1]  -replace '\D+(\d+)','$1')
        $true
      }
      catch{
        $false
      }
    }
    
  }
}

# Connect to Sharepoint Online
function Connect-SharePointOnlinePS {
  param (
    $Creds=$MyCred,
    [Parameter(Mandatory=$true)]$Tennant
  )
  Import-Module -Name "$env:ProgramFiles\SharePoint Online Management Shell\Microsoft.Online.SharePoint.PowerShell\Microsoft.Online.SharePoint.PowerShell.psd1" -WarningAction SilentlyContinue
  # Clean up existing PowerShell Sessions Sharepoint Online
  try{
    Disconnect-SPOService -ErrorAction SilentlyContinue
  }
  catch{}
  
  # Connect to Sharepoint Online
  Connect-SPOService -Url https://$Tennant-admin.sharepoint.com -Credential $Creds
  Write-Output -InputObject 'Connected to SharePoint Online PowerShell'
}

# Connect to Msol
function Connect-MsolPS {
  [CmdletBinding()]
  param (
    $Creds=$MyCred
  )
  # Import required modules
  Import-Module -Name MSOnline
  # Connect to Azure AD
  Connect-MsolService -Credential $Creds |Out-Null
  Write-Output -InputObject 'Connected to MSOL PowerShell'
}

# Connect to Compliance PowerShell
function Connect-CompliancePS {
  [CmdletBinding()]
  param (
    $Creds=$MyCred
  )
  # Clean up existing PowerShell Sessions for Exchange Online
  (Get-PSSession).where({$_.ComputerName -like '*compliance.protection.outlook.com'}) | Remove-PSSession
  # Connect to Exchange Online
  $ComplianceSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.compliance.protection.outlook.com/powershell-liveid/ -Credential $Creds -Authentication Basic -AllowRedirection -WarningAction SilentlyContinue
  $null = Import-PSSession -Session $ComplianceSession -DisableNameChecking:$true -AllowClobber:$true 
  Write-Output -InputObject 'Connected to Compliance and Protection'
}

# Get credential's from assets
$Credentials = Get-AutomationPSCredential -Name 'Office365 Service Account 01'

# Connect to necessary services
Write-Output 'Connecting to Azure AD'
Connect-MsolPS -Creds $Credentials
Write-Output 'Connecting to Sharepoint Online'
Connect-SharePointOnlinePS -Creds $Credentials -Tennant hennesandmauritz
Write-Output 'Connecting to Compliance and Protection'
Connect-CompliancePS -Creds $Credentials

# Get all OneDrive sites and pipe to function Select-SharePointSite
Write-Output 'Get all OneDrive sites'
$SPOSites = Get-SPOSite -Template 'SPSPERS#10' -IncludePersonalSite:$true -Limit All
Write-Output 'Process sites with function Select-SharePointSite'
$SPOSites = $SPOSites | Select-SharePointSite

# Get Specific Swiss users (Switzerland)
Write-Output -InputObject 'Collect Specific Swiss Users in scope'
$CHAccount = Get-MsolGroupMember -GroupObjectId '47c34d2f-7850-4ed1-bc43-43d38a0487d6' -All
$CHPersonal = Get-MsolGroupMember -GroupObjectId '21282a0a-619f-4f50-974a-dedb9533a839' -All
$CHHR = Get-MsolGroupMember -GroupObjectId 'd373191f-21d7-460a-ab92-d4a319238ed9' -All
$CHSecurity = Get-MsolGroupMember -GroupObjectId '38d3a549-b5e3-4daf-8fdc-087f83ac28be' -All
# Gather all Swiss in one object
$Swissv0 = $CHAccount + $CHPersonal + $CHHR + $CHSecurity
$CHUsers = $Swissv0 | Get-MsolUser | Select-MSOLUser

# Get Specific US users (United states)
Write-Output -InputObject 'Collect Specific US users in scope'
$USUsers = Get-MsolUser -UsageLocation US -all
# Get Specific CA users (Canada)
Write-Output -InputObject 'Collect Specific CA users in scope'
$CAUsers = Get-MsolUser -UsageLocation CA -all
# Combine US and CA to NA
$NAUsers = $CAUsers + $USUsers | Select-MSOLUser

# Combinding all msol users in-scope that are licensed to one object
Write-Output -InputObject 'Combining US and CA users to one object and filtering out only licensed users'
$UsersInScope = $NAUsers.Where({$_.IsLicensed -eq $true}) + $CHUsers.Where({$_.IsLicensed -eq $true})

# Getting different policys
# 7 years
$odfbpolicy7yname = '7 Years Retention for SPO & OD4B'
$odfbpolicy7ynameprefix = $odfbpolicy7yname
$message = 'Collecting Compliance policy members of: '
$outputmessage = $message + $odfbpolicy7yname
Write-Output $outputmessage
$odfbpolicy7y = Get-RetentionCompliancePolicy -DistributionDetail | Where-Object { $_.Name -like "$odfbpolicy7yname*" }
$message = "Selecting latest $odfbpolicy7yname policy"
Write-Output $message
$latestodfbpolicy7y = ($odfbpolicy7y | Sort-Object WhenCreated -Descending | Select-Object -First 1).name
$message = "Latest policy selected $($latestodfbpolicy7y)"
Write-Output $message
$odfbpolicy7yusers = $odfbpolicy7y.OneDriveLocation | Select-PolicyUser
$odfbpolicy7ypreservationduration = 2555
  
# 10 years
$odfbpolicy10yname = '10 Years Retention for SPO & OD4B'
$odfbpolicy10ynameprefix = $odfbpolicy10yname
$message = 'Collecting Compliance policy members of: '
$outputmessage = $message + $odfbpolicy10yname
Write-Output $outputmessage
$odfbpolicy10y = Get-RetentionCompliancePolicy -DistributionDetail | Where-Object { $_.Name -like "$odfbpolicy10yname*" }
$message = "Selecting latest $odfbpolicy10yname policy"
Write-Output $message
$latestodfbpolicy10y = ($odfbpolicy10y | Sort-Object WhenCreated -Descending | Select-Object -First 1).name
$message = "Latest policy selected $($latestodfbpolicy10y)"
Write-Output $message
$odfbpolicy10yusers = $odfbpolicy10y.OneDriveLocation | Select-PolicyUser
$odfbpolicy10ypreservationduration = 3650
  
# Define the way the retentionpolicy calculates expiration
$ExpirationDateOption = 'ModificationAgeInDays'
  
# Looping each user in scope, constructing SPO Url and checking policy membership
Write-Output -InputObject 'Start processing each user in scope'
foreach ($user in $UsersInScope){

  $UPN = $user.UserPrincipalName
  $SPOURL = 'https://hennesandmauritz-my.sharepoint.com/personal/'
  $URLGuess = ($SPOURL+$user.UserPrincipalName.Replace('.','_').replace('@','_')).ToLower()
    
  if ($URLGuess -in $SPOSites.url){
    # OneDrive site was found for the user
    Write-Output -InputObject "ODFB site $urlguess found for: $UPN"

    # Where is the user located
    if ($upn -in $NAUsers.userprincipalname){
      # user was found in NA variable
      Write-Output -InputObject 'ODFB site belongs to NA'
      # check if user already have policy enabled
      if ($URLGuess -in $odfbpolicy7yusers.site){
        Write-Output -InputObject "ODFB site already has the following policy applied: $odfbpolicy7ynameprefix"
      }
      else{
        # Policy needs to be applied
        Write-Output -InputObject "Applying policy: $latestodfbpolicy7y to $urlguess"
        if ($ProductionMode -eq $true){
          try{
            Set-RetentionCompliancePolicy -Identity $latestodfbpolicy7y -AddOneDriveLocation $URLGuess
          }
          catch{
            if($Error[0].Exception.Message -like '*Please select fewer sites to continue*')
            {
              Write-Output -InputObject "Policy: $latestodfbpolicy7y is at max capacity, creating new policy"
              $NewPolicyNum = 1 + ($latestodfbpolicy7y -replace "$($odfbpolicy7ynameprefix+'_')")
              $latestodfbpolicy7y = "$odfbpolicy7ynameprefix"+'_'+"$NewPolicyNum"
         
              try
              {
                Write-Output -InputObject "New policy: $latestodfbpolicy7y"
                New-RetentionCompliancePolicy -Name $latestodfbpolicy7y | Out-Null
                Write-Output -InputObject "Creating new policy compliance rule"
                New-RetentionComplianceRule -Name $latestodfbpolicy7y -Policy $latestodfbpolicy7y -RetentionDuration $odfbpolicy7ypreservationduration -ExpirationDateOption $ExpirationDateOption | Out-Null
                Write-Output -InputObject "Now applying new policy: $latestodfbpolicy7y to $urlguess"
                Set-RetentionCompliancePolicy -Identity $latestodfbpolicy7y -AddOneDriveLocation $URLGuess | Out-Null
              }
              catch
              {
                Write-Output -InputObject "Failed when creating $latestodfbpolicy7y. Error: $($Error[0].Exception.message)"
                Exit
              }
                   
            }
            else{
              Write-Output -InputObject "Failed to apply policy $latestodfbpolicy7y to $URLGuess. Error: $($Error[0].Exception.message)"
            }
          }
        }
      }
    }
      
    # If user is not in US or CA then the user MUST be in CH
    elseif ($upn -in $CHUsers.userprincipalname) {
      # User in CH
      Write-Output -InputObject "ODFB site $urlguess belongs to CH"
      if ($URLGuess -in $odfbpolicy10yusers.site){
        Write-Output -InputObject "ODFB site already has the following policy applied: $odfbpolicy10ynameprefix"
      }
      else{
        # Policy needs to be applied
        Write-Output -InputObject "Applying policy: $latestodfbpolicy10y to $urlguess"
        if ($ProductionMode -eq $True){
          try{
            Set-RetentionCompliancePolicy -Identity $latestodfbpolicy10y -AddOneDriveLocation $URLGuess
          }
          catch{
            if($Error[0].Exception.Message -like '*Please select fewer sites to continue*')
            {
              Write-Output -InputObject "Policy: $latestodfbpolicy10y is at max capacity, creating new policy"
              $NewPolicyNum = 1 + ($latestodfbpolicy10y -replace "$($odfbpolicy10ynameprefix+'_')")
              $latestodfbpolicy10y = "$odfbpolicy10ynameprefix"+'_'+"$NewPolicyNum"
         
              try
              {
                Write-Output -InputObject "New policy: $latestodfbpolicy10y"
                New-RetentionCompliancePolicy -Name $latestodfbpolicy10y | Out-Null
                Write-Output -InputObject "Creating new policy compliance rule"
                New-RetentionComplianceRule -Name $latestodfbpolicy10y -Policy $latestodfbpolicy10y -RetentionDuration $odfbpolicy10ypreservationduration -ExpirationDateOption $ExpirationDateOption | Out-Null
                Write-Output -InputObject "Now applying new policy: $latestodfbpolicy10y to $urlguess"
                Set-RetentionCompliancePolicy -Identity $latestodfbpolicy10y -AddOneDriveLocation $URLGuess | Out-Null
              }
              catch
              {
                Write-Output -InputObject "Failed when creating $latestodfbpolicy10y. Error: $($Error[0].Exception.message)"
                Exit
              }
            }
            else{
              Write-Output -InputObject "Failed to apply policy $latestodfbpolicy10y to $URLGuess. Error: $($Error[0].Exception.message)"
            }
          }
        }
      }
    }
  }
  else{
    # User does not have SPOsite
  }
}

# Looping each SPO Url ending with numbers and checking policy membership
Write-Output -InputObject 'Start processing each site ending with numbers'
$OneDriveSitesEndingInNumbers = $SPOSites.Where({$_.SiteWithNumber[1] -eq 'True'})
foreach ($OneDriveSite in $OneDriveSitesEndingInNumbers){
  
  $EndURLWithoutNumber = $OneDriveSite.url.Split('_')[-1] -replace $($OneDriveSite.SiteWithNumber[0])
  $EndURLNumber = $OneDriveSite.url.Split('_')[-1]
  $UpnFromUrl = $OneDriveSite.Url.Split('/')[-1]
  $DomainFromUrl = $UpnFromUrl.Split('_')[-2]
  # Building UPN 
  $guessUPN = $null
  foreach ($obj in $UpnFromUrl.Split('_')){
    if ($obj -eq $EndURLNumber){
      $obj = "$EndURLWithoutNumber"
      $guessUPN += $obj
    }
    elseif ($obj -eq $DomainFromUrl){
      $obj = '@'+"$DomainFromUrl"
      $guessUPN += $obj+'.'
    }
    else{
      $guessUPN += $obj+'.'
    }
  }
  $guessUPN = $guessUPN.Replace('.@','@')
    
  #$UsersInScope.ProxyAddresses  
  if ($guessUPN -in $UsersInScope.ProxyAddressesWithoutSMTP ){
    # Guessed UPN is in one of the PolicyGroups
    
    # Where is the user located
    if ($guessUPN -in $NAUsers.ProxyAddressesWithoutSMTP){
      # user was found in NA variable
      Write-Output -InputObject "ODFB site $($OneDriveSite.url) found for: $guessUPN"
      # check if user already have policy enabled
      if ($OneDriveSite.Url -in $odfbpolicy7yusers.site){
        Write-Output -InputObject "ODFB site already has the following policy applied: $odfbpolicy7ynameprefix"
      }
      else{
        # Policy needs to be applied
        Write-Output -InputObject "Applying policy: $latestodfbpolicy7y to $($OneDriveSite.url)"
        if ($ProductionMode -eq $True){
          try{
            Set-RetentionCompliancePolicy -Identity $latestodfbpolicy7y -AddOneDriveLocation $OneDriveSite.Url
          }
          catch{
            if($Error[0].Exception.Message -like '*Please select fewer sites to continue*')
            {
              Write-Output -InputObject "Policy: $latestodfbpolicy7y is at max capacity, creating new policy"
              $NewPolicyNum = 1 + ($latestodfbpolicy7y -replace "$($odfbpolicy7ynameprefix+'_')")
              $latestodfbpolicy7y = "$odfbpolicy7ynameprefix"+'_'+"$NewPolicyNum"
              try
              {
                Write-Output -InputObject "New policy: $latestodfbpolicy7y"
                New-RetentionCompliancePolicy -Name $latestodfbpolicy7y | Out-Null
                Write-Output -InputObject "Creating new policy compliance rule"
                New-RetentionComplianceRule -Name $latestodfbpolicy7y -Policy $latestodfbpolicy7y -RetentionDuration $odfbpolicy7ypreservationduration -ExpirationDateOption $ExpirationDateOption | Out-Null
                Write-Output -InputObject "Now applying new policy: $latestodfbpolicy7y to $($OneDriveSite.url)" 
                Set-RetentionCompliancePolicy -Identity $latestodfbpolicy7y -AddOneDriveLocation $OneDriveSite.Url | Out-Null
              }
              catch
              {
                Write-Output -InputObject "Failed when creating $latestodfbpolicy7y. Error: $($Error[0].Exception.message)"
                Exit
              }
            }
            else{
              Write-Output -InputObject "Failed to apply policy $latestodfbpolicy7y to $($OneDriveSite.Url). Error: $($Error[0].Exception.message)"
            }
          }
        }
      }
    }
      
    # User is not in NA then the user must be in CH
    elseif ($guessUPN -in $CHUsers.ProxyAddressesWithoutSMTP) {
      # User in CH
      Write-Output -InputObject "ODFB site $($OneDriveSite.url) found for: $guessUPN"
      if ($OneDriveSite.Url -in $odfbpolicy10yusers.site){
        Write-Output -InputObject "ODFB site already has the following policy applied: $odfbpolicy10ynameprefix"
      }
      else{
        # Policy needs to be applied
        Write-Output -InputObject "Applying policy: $latestodfbpolicy10y to $($OneDriveSite.url)"
        if ($ProductionMode -eq $True){
          try{
            Set-RetentionCompliancePolicy -Identity $latestodfbpolicy10y -AddOneDriveLocation $OneDriveSite.url
          }
          catch{
            if($Error[0].Exception.Message -like '*Please select fewer sites to continue*')
            {
              Write-Output -InputObject "Policy: $latestodfbpolicy10y is at max capacity, creating new policy"
              $NewPolicyNum = 1 + ($latestodfbpolicy10y -replace "$($odfbpolicy10ynameprefix+'_')")
              $latestodfbpolicy10y = "$odfbpolicy10ynameprefix"+'_'+"$NewPolicyNum"
         
              try
              {
                Write-Output -InputObject "New policy: $latestodfbpolicy10y"
                New-RetentionCompliancePolicy -Name $latestodfbpolicy10y | Out-Null
                Write-Output -InputObject "Creating new policy compliance rule"
                New-RetentionComplianceRule -Name $latestodfbpolicy10y -Policy $latestodfbpolicy10y -RetentionDuration $odfbpolicy10ypreservationduration -ExpirationDateOption $ExpirationDateOption | Out-Null
                Write-Output -InputObject "Now applying new policy: $latestodfbpolicy10y to $($OneDriveSite.url)"
                Set-RetentionCompliancePolicy -Identity $latestodfbpolicy10y -AddOneDriveLocation $OneDriveSite.url | Out-Null
              }
              catch
              {
                Write-Output -InputObject "Failed when creating $latestodfbpolicy10y. Error: $($Error[0].Exception.message)"
                Exit
              }
            }
            else{
              Write-Output -InputObject "Failed to apply policy $latestodfbpolicy10y to $($OneDriveSite.url). Error: $($Error[0].Exception.message)"
            }
          }
        }
      }
    }
  }
}

# Get all MSOL Domains for further troubleshooting
$MSOLDOMAINS = Get-MsolDomain
# Looping each problem domain in $MSOLDOMAINS that contains a subdomain
Write-Output -InputObject 'Start processing each site ending with a subdomain'
foreach ($problemdomain in $($MSOLDOMAINS.Where({$_.Name -like '*.*.*'}))){

  $problemDomainWithUnderLine = $problemdomain.Name.Replace('.','_')
  foreach ($problemuser in $($SPOSites.Where({$_.url -like "*$problemDomainWithUnderLine"}))){
    if ($(($problemuser.Url).EndsWith("_$problemDomainWithUnderLine")) -eq $true){
      
      $PreFix = ($problemuser.url.Replace($problemDomainWithUnderLine,'').split('/')[-1]).replace('_','.')
      $sufix = $problemDomainWithUnderLine.Replace('_','.')
      $Upn = ($PreFix+'@'+$sufix).Replace('.@','@')
      
      if ($UPN -in $UsersInScope.ProxyAddressesWithoutSMTP ){
        # Guessed UPN is in one of the PolicyGroups

        # Where is the user located
        if ($UPN -in $NAUsers.ProxyAddressesWithoutSMTP){
          # user was found in NA variable
          Write-Output -InputObject "ODFB site $($problemuser.url) found for: $UPN"
          # check if user already have policy enabled
          if ($problemuser.Url -in $odfbpolicy7yusers.site){
            Write-Output -InputObject "ODFB site already has the following policy applied: $odfbpolicy7ynameprefix"
          }
          else{
            # Policy needs to be applied
            Write-Output -InputObject "Applying policy: $latestodfbpolicy7y to $($problemuser.url)"
            if ($ProductionMode -eq $True){
              try{
                Set-RetentionCompliancePolicy -Identity $latestodfbpolicy7y -AddOneDriveLocation $problemuser.url
              }
              catch{
                if($Error[0].Exception.Message -like '*Please select fewer sites to continue*')
                {
                  Write-Output -InputObject "Policy: $latestodfbpolicy7y is at max capacity, creating new policy"
                  $NewPolicyNum = 1 + ($latestodfbpolicy7y -replace "$($odfbpolicy7ynameprefix+'_')")
                  $latestodfbpolicy7y = "$odfbpolicy7ynameprefix"+'_'+"$NewPolicyNum"
         
                  try
                  {
                    Write-Output -InputObject "New policy: $latestodfbpolicy7y"
                    New-RetentionCompliancePolicy -Name $latestodfbpolicy7y | Out-Null
                    Write-Output -InputObject "Creating new policy compliance rule"
                    New-RetentionComplianceRule -Name $latestodfbpolicy7y -Policy $latestodfbpolicy7y -RetentionDuration $odfbpolicy7ypreservationduration -ExpirationDateOption $ExpirationDateOption | Out-Null
                    Write-Output -InputObject "Now applying new policy: $latestodfbpolicy7y to $($problemuser.url)"
                    Set-RetentionCompliancePolicy -Identity $latestodfbpolicy7y -AddOneDriveLocation $problemuser.url | Out-Null
                  }
                  catch
                  {
                    Write-Output -InputObject "Failed when creating $latestodfbpolicy7y. Error: $($Error[0].Exception.message)"
                    Exit
                  }
                   
                }
                else{
                  Write-Output -InputObject "Failed to apply policy $latestodfbpolicy7y to $($problemuser.url). Error: $($Error[0].Exception.message)"
                }
              }
            }
          }
        }
      
        # User is not in NA then the user must be in CH
        elseif ($UPN -in $CHUsers.ProxyAddressesWithoutSMTP) {
          # User in CH
          Write-Output -InputObject "ODFB site $($problemuser.url) found for: $UPN"
          if ($problemuser.Url -in $odfbpolicy10yusers.site){
            Write-Output -InputObject "ODFB site already has the following policy applied: $odfbpolicy10ynameprefix"
          }
          else{
            # Policy needs to be applied
            Write-Output -InputObject "Applying policy: $latestodfbpolicy10y to $($problemuser.url)"
            if ($ProductionMode -eq $True){
              try{
                Set-RetentionCompliancePolicy -Identity $latestodfbpolicy10y -AddOneDriveLocation $problemuser.url
              }
              catch{
                if($Error[0].Exception.Message -like '*Please select fewer sites to continue*')
                {
                  Write-Output -InputObject "Policy: $latestodfbpolicy10y is at max capacity, creating new policy"
                  $NewPolicyNum = 1 + ($latestodfbpolicy10y -replace "$($odfbpolicy10ynameprefix+'_')")
                  $latestodfbpolicy10y = "$odfbpolicy10ynameprefix"+'_'+"$NewPolicyNum"
                  try
                  {
                    Write-Output -InputObject "New policy: $latestodfbpolicy10y"
                    New-RetentionCompliancePolicy -Name $latestodfbpolicy10y | Out-Null
                    Write-Output -InputObject "Creating new policy compliance rule"
                    New-RetentionComplianceRule -Name $latestodfbpolicy10y -Policy $latestodfbpolicy10y -RetentionDuration $odfbpolicy10ypreservationduration -ExpirationDateOption $ExpirationDateOption | Out-Null
                    Write-Output -InputObject "Now applying new policy: $latestodfbpolicy10y to $($problemuser.url)"
                    Set-RetentionCompliancePolicy -Identity $latestodfbpolicy10y -AddOneDriveLocation $problemuser.url | Out-Null
                  }
                  catch
                  {
                    Write-Output -InputObject "Failed when creating $latestodfbpolicy10y. Error: $($Error[0].Exception.message)"
                    Exit
                  }
                }
                else{
                  Write-Output -InputObject "Failed to apply policy $latestodfbpolicy10y to $($problemuser.url). Error: $($Error[0].Exception.message)"
                }
              }
            }
          }
        }
      }
    }
  }
}
Write-Output -InputObject 'Processing done'