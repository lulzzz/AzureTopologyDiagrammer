Function Get-ResourceGroups {

    $BaseUri = "https://management.azure.com/subscriptions/$($Script:Currentsubscription.SubscriptionId)/resourcegroups" 


    $AuthToken = $Script:Currentsubscription.AccessToken
    $TokenExpirationUtc = $Script:Currentsubscription.Expiry

    $headers = @{"Authorization"="Bearer $AuthToken"
                    "x-ms-version" = "2013-08-01";
                    "Content-Type"="application/json"}

    $Uri = "https://management.azure.com/subscriptions/$($Script:Currentsubscription.SubscriptionId)/resourcegroups?api-version=2016-09-01" 

    $ResourceGroups = (Invoke-RestMethod -Method Get -Uri $Uri -Headers $headers).value

    return $ResourceGroup
}


Export-ModuleMember Get-ResourceGroups