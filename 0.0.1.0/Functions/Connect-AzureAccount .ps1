Function Connect-AzureAccount {

    $Script:AuthResult = Get-AuthenticationResult -url $Script:LoginUrl
    
    # Set up our HTTP headers
    $AuthToken =  $Script:AuthResult.AccessToken
    $RefreshToken =  $Script:AuthResult.RefreshToken
    $TokenExpirationUtc =  $Script:AuthResult.ExpiresOn

    $headers = @{
        "Authorization"="Bearer $AuthToken"
        "Content-Type"="application/json"
    }

    # List all tenants at first
    $tenants = (Invoke-RestMethod -Method Get -Uri "https://management.azure.com/tenants?api-version=2016-09-01" -Headers $headers).value

    $Script:AllSubscriptions = @()
    
    Foreach ($tenant in  $tenants) {

        $LoginURI = "https://login.windows.net/$($Tenant.tenantId)/oauth2/authorize/"

        $Tenantauth = Get-AuthenticationResult -url $LoginUrI

        $AuthToken = $Tenantauth.AccessToken
        $TokenExpirationUtc = $Tenantauth.ExpiresOn

        $headers = @{"Authorization"="Bearer $AuthToken"
                    "x-ms-version" = "2013-08-01";
                    "Content-Type"="application/json"}
        
        $SubscriptionResult = Invoke-RestMethod -Method Get -Uri "https://management.azure.com/subscriptions?api-version=2016-09-01" -Headers $headers

        foreach ($Subscription in $SubscriptionResult.Value)  {

                $Script:AllSubscriptions += [PSCUstomObject]@{"SubscriptionId" = $Subscription.subscriptionId
                                                            "DisplayName" = $Subscription.displayName
                                                            "State" = $Subscription.state
                                                            "AccessToken" = $Tenantauth.AccessToken
                                                            "RefreshToken" = $Tenantauth.RefreshToken
                                                            "Expiry" = $Tenantauth.ExpiresOn
                                                            "TenantId" = $Tenant.tenantId}

        }
    }

    if ($Script:AllSubscriptions.count -eq 1) {
        $Script:Currentsubscription = $Script:AllSubscription[0]
    }
    elseif ($Script:AllSubscriptions.count -gt 1) {
        Write-Warning -Message "Multiple Subscriptions found and none specified. Please select the desired one"
        $i = 0
        $list = foreach ($T in $Script:AllSubscriptions) {
            $i++
            '[{0}] - {1} - {2}' -f $i,$T.DisplayName,$T.SubscriptionId
        }
        do {
            $strResult = Read-Host -Prompt "Enter the index number of the desired subscription: `n$($List | Out-String)"
            try {
                [int]$Result = [convert]::ToInt32($strResult, 10)    
            }
            catch {
                $Result = 0
            }
            
        } while ($result -gt $Script:AllSubscriptions.count -or $result -eq 0)

        $Script:Currentsubscription = $Script:AllSubscriptions[$result -1]

    }
    elseif ($Script:AllSubscriptions.count -eq 0) {
        Write-Error -Message "Can't get any subscription for this account"
    }
}

Export-ModuleMember Connect-AzureAccount