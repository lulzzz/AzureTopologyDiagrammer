Function Get-AuthenticationResult {

    param (
        $url
    )
   
   Try {
        $PromptBehavior = [Microsoft.IdentityModel.Clients.ActiveDirectory.PromptBehavior]::Auto

        $AuthContext = New-Object -TypeName Microsoft.IdentityModel.Clients.ActiveDirectory.AuthenticationContext -ArgumentList ($url)
        Write-Output "$($Script:ResourceUrl) $($Script:ClientId)  $($Script:RedirectUri) $($PromptBehavior)"
        $Script:authResult = $AuthContext.AcquireToken($Script:ResourceUrl,$Script:ClientId, $Script:RedirectUri, $PromptBehavior)

        if ($Script:authResult) {
        # Return our auth result
            return $Script:authResult
        }
        else {
            Write-Error "Error Authenticate"
        }
   }
   Catch {
       throw
   }
    
}

Export-ModuleMember Get-AuthenticationResult