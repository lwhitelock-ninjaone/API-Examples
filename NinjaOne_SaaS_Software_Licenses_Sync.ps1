# Creation Mode
#$Mode = 'CUSTOM'
$Mode = 'ENDUSER'

$CreateAccountsIfNonExisting = $True
# As users can only be unique globaly across an instance in NinjaOne provide an option to insert a + address section to email addresses, so that we don't run into user conflict problems
# For example change Luke@example.com to Luke+LukeTesting@example.com
# Has two options as Powershell doesn't like doing a replace on strings starting with +
$PlusAddressRemoveOption = '\+LukeTesting'
$PlusAddressInsertOption = '+LukeTesting'

# Ninja SaaS API Details
$APIURL = 'https://na-saas-npp.backup.ninjarmm.com/api/'
$ResellerToken = 'ResellerToken'
$AuthenticationToken = 'AuthToken'

$SaaSAuthHeader = @{
    'X-Reseller-Token' = $ResellerToken
    'X-Access-Token'   = $AuthenticationToken
}


# NinjaOne API Details
$ClientID = 'ClientID'
$Secret = 'Secret'
$NinjaInstance = 'app.ninjarmm.com'


$AuthBody = @{
    'grant_type'    = 'client_credentials'
    'client_id'     = $ClientID
    'client_secret' = $Secret
    'scope'         = 'monitoring management' 
}



$Result = Invoke-WebRequest -uri "https://$($NinjaInstance)/ws/oauth/token" -Method POST -Body $AuthBody -ContentType 'application/x-www-form-urlencoded'

$NinjaAuthHeader = @{
    'Authorization' = "Bearer $(($Result.content | ConvertFrom-Json).access_token)"
}


# Fetch Dropsuite Plans and Users
$Users = (Invoke-WebRequest -Uri "$($APIURL)/users" -Method Get -Headers $SaaSAuthHeader).content | ConvertFrom-Json -Depth 100
$Plans = (Invoke-WebRequest -Uri "$($APIURL)/plans" -Method Get -Headers $SaaSAuthHeader).content | ConvertFrom-Json -Depth 100

# For syncing just the license counts in custom mode:
if ($Mode -eq 'CUSTOM') {
    Write-Host 'Starting Custom License Mode'
    # Loop each dropsuite user
    foreach ($SaaSOrganization in $Users) {
        # Find the plan for the user
        $Plan = $Plans | Where-Object { $_.id -eq $SaaSOrganization.plan_id }
    
        # Get current license counts
        $LicensesInUse = $SaaSOrganization.seats_used
        $LicensesAvailable = $SaaSOrganization.seats_available

        # Create the body for the upsert endpoint. This will match to an existing license or create a new one
        $CreateUpdateLicense = @{
            name          = "$($SaaSOrganization.organization_name) - $($Plan.name)"
            description   = ''
            type          = 'CUSTOM'
            publisherName = 'NinjaOne SaaS Backup'
            vendorName    = 'NinjaOne'
            scope         = @{
                organizationNames = @($($SaaSOrganization.organization_name))
            }
            quantity      = $LicensesAvailable
            currentUsage  = $LicensesInUse
        } | ConvertTo-Json

        $Result = (Invoke-WebRequest -uri "https://$($NinjaInstance)/api/v2/software-license/upsert" -Method POST -Headers $NinjaAuthHeader -Body $CreateUpdateLicense -ContentType 'application/json').content | ConvertFrom-Json -depth 100

    }

    # For mode = End User
} ElseIf ($Mode -eq 'ENDUSER') {
    Write-Host 'Starting End User License Mode'
    # Get all users from NinjaOne
    $NinjaExistingUsers = (Invoke-WebRequest -uri "https://$($NinjaInstance)/api/v2/users" -Method GET -Headers $NinjaAuthHeader -ContentType 'application/json').content | ConvertFrom-Json -depth 100
    
    # Strip out the plus address for matching on original email
    $NinjaExistingUsers | ForEach-Object {
        $_.email = $_.email -replace $PlusAddressRemoveOption, ''
    }

    # Get all organizations from NinjaOne
    $NinjaOrganizations = (Invoke-WebRequest -uri "https://$($NinjaInstance)/api/v2/organizations" -Method GET -Headers $NinjaAuthHeader -ContentType 'application/json').content | ConvertFrom-Json -depth 100

    foreach ($SaaSOrganization in $Users) {
        # Get User auth token for Dropsuite
        $UserSaaSAuthHeader = @{
            'X-Reseller-Token' = $ResellerToken
            'X-Access-Token'   = $SaaSOrganization.Authentication_Token
        }

        $MatchedNinjaOneOrganization = $NinjaOrganizations | Where-Object { $_.name -eq $SaasOrganization.organization_name }
        if (($MatchedNinjaOneOrganization | Measure-Object).count -eq 1) {

            # Find the plan for the user
            $Plan = $Plans | Where-Object { $_.id -eq $SaaSOrganization.plan_id }
        
            # Get accounts for the user
            $Accounts = (Invoke-WebRequest -Uri "$($APIURL)/accounts" -Method Get -Headers $UserSaaSAuthHeader).content | ConvertFrom-Json -Depth 100

            $LicensedUsers = Foreach ($UserAccount in $Accounts) {
                # Check if the user exists in Ninja
                $MatchedUser = $NinjaExistingUsers | Where-Object {$_.email -eq $UserAccount.email}
                $MatchedCount = ($MatchedUser | Measure-Object).count
                
                # Split apart the email address at the @ symbol
                $UserParts = ($UserAccount.email) -split ('@')
                # Insert the +address if configured
                $ParsedAddress = "$($UserParts[0])$($PlusAddressInsertOption)@$($UserParts[1])"
            
                # Handle existing users.
                if ($MatchedCount -eq 1) {
                    Write-Host "$($UserAccount.email) already exists in Ninja"
                    $ParsedAddress
                } elseif ($MatchedCount -eq 0) {
                    # Handle new users.
                    $NameParts = ($UserParts[0] -split '`.')
                    $FirstName = $NameParts[0]
                    If ($NameParts[1]) {
                        $LastName = $NameParts[1]
                    } else {
                        $LastName = $SaasOrganization.organization_name
                    }
                    if ($CreateAccountsIfNonExisting -eq $True) {
                        $UserCreate = @{
                            firstName      = $FirstName
                            lastName       = $LastName
                            email          = $ParsedAddress
                            organizationId = $MatchedNinjaOneOrganization.id
                        } | ConvertTo-Json
                        try {
                            $Result = (Invoke-WebRequest -uri "https://$($NinjaInstance)/api/v2/user/end-users" -Method POST -Headers $NinjaAuthHeader -Body $UserCreate -ContentType 'application/json' -ea stop).content | ConvertFrom-Json -depth 100
                            $ParsedAddress
                        } catch {
                            Write-Error "Failed to create user $($ParsedAddress): $_"
                        }
                    }
                }
            }

            # Get current license counts
            $LicensesAvailable = $SaaSOrganization.seats_available
    
            # Create the body for the upsert endpoint. This will match to an existing license or create a new one
            $CreateUpdateLicense = @{
                name             = "$($SaaSOrganization.organization_name) - $($Plan.name)"
                description      = ''
                type             = 'PER_USER'
                publisherName    = 'NinjaOne SaaS Backup'
                vendorName       = 'NinjaOne'
                scope            = @{
                    organizationNames = @($($SaaSOrganization.organization_name))
                }
                quantity         = $LicensesAvailable
                currentLicensees = $LicensedUsers
            } | ConvertTo-Json
    
            $Result = (Invoke-WebRequest -uri "https://$($NinjaInstance)/api/v2/software-license/upsert" -Method POST -Headers $NinjaAuthHeader -Body $CreateUpdateLicense -ContentType 'application/json').content | ConvertFrom-Json -depth 100
        } else {
            Write-Error "$($SaaSOrganization.organization_name) could not be matched to an existing organization in NinjaOne"
        }
    
    } 
}

