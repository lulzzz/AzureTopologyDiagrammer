# ===================================================================================
# Name: Azure Topology Diagrammer
# Desc: Functions to query Azure REST API
# ===================================================================================

Function Get-ResourceGroups
{
    $queryResult = Execute-ResourceManagerQuery -ApiAction Get -QueryBase "resourceGroups"
    return $queryResult
}

Function Load-ActiveDirectoryAuthenticationLibrary
{
    # Adapted from the excellent work here:
    #   http://www.dushyantgill.com/blog/2013/12/27/aadgraphpowershell-a-powershell-client-for-windows-azure-ad-graph-api/
    #   https://github.com/dushyantgill/AADGraphPowerShell/blob/master/AADGraph.psm1

    $nugetBinaryPath = "$pwd\Nuget.exe"
    $nugetUrl = "http://www.nuget.org/nuget.exe"
    $nugetPackagesPath = "$pwd\Nugets"    

    # Check for Nuget packages directory
    if (!(Test-Path -Path $nugetPackagesPath))
    {
        # Create the directory for Nuget packages
        New-Item -ItemType Directory -Path $nugetPackagesPath        
    }

    # Check to see if ADAL available
    $adalPackageDirectories = Get-ChildItem -Path $nugetPackagesPath -Filter "Microsoft.IdentityModel.Clients.ActiveDirectory*" -Directory

    # Download binaries if ADAL not available
    if($adalPackageDirectories.Length -eq 0)
    {
        # Check to see if we've already downloaded Nuget
        if (!(Test-Path -Path $nugetBinaryPath))
        {
            # If Nuget not available, download using WebClient
            Write-Host "nuget.exe not found. Downloading from $nugetUrl ..." -ForegroundColor Yellow
            $wc = New-Object System.Net.WebClient
            $wc.DownloadFile($nugetUrl,$nugetBinaryPath);
        }

        # Download ADAL
        $nugetDownloadExpression = "$nugetBinaryPath install Microsoft.IdentityModel.Clients.ActiveDirectory -OutputDirectory $nugetPackagesPath | Out-Null"
        Invoke-Expression $nugetDownloadExpression

        # Retrieve the path of our newly downloaded library
        $adalPackageDirectories = Get-ChildItem -Path $nugetPackagesPath -Filter "Microsoft.IdentityModel.Clients.ActiveDirectory*" -Directory
    }
    
    # Get references to the DLLs (only including net45 binaries)
    $adalAssembly = Get-ChildItem "Microsoft.IdentityModel.Clients.ActiveDirectory.dll" -Path "$($adalPackageDirectories[$adalPackageDirectories.length-1].FullName)\lib\net45" -Recurse
    $adalWinFormsAssembly = Get-ChildItem "Microsoft.IdentityModel.Clients.ActiveDirectory.WindowsForms.dll" -Path "$($adalPackageDirectories[$adalPackageDirectories.length-1].FullName)\lib\net45" -Recurse

    # Load the DLLs via Reflection
    if($adalAssembly.Length -gt 0 -and $adalWinFormsAssembly.Length -gt 0)
    {
        Write-Host "Loading ADAL Assemblies ..." -ForegroundColor Green
        [System.Reflection.Assembly]::LoadFrom($adalAssembly.FullName) | Out-Null
        [System.Reflection.Assembly]::LoadFrom($adalWinFormsAssembly.FullName) | Out-Null
        return $true
    }
    else{
        Write-Host "Fixing Active Directory Authentication Library package directories ..." -ForegroundColor Yellow
        $adalPackageDirectories | Remove-Item -Recurse -Force | Out-Null
        Write-Host "Not able to load ADAL assembly. Delete the Nugets subfolder, restart your PowerShell session and try again ..."
        return $false
    }
}

Function Get-AuthenticationResult($tenant = "common", $env="prod")
{
    # Adapted from the excellent work here:
    #   http://www.dushyantgill.com/blog/2013/12/27/aadgraphpowershell-a-powershell-client-for-windows-azure-ad-graph-api/
    #   https://github.com/dushyantgill/AADGraphPowerShell/blob/master/AADGraph.psm1

    $loadedAssemblies = [System.AppDomain]::CurrentDomain.GetAssemblies() | Where-Object {$_.Location -like "*Microsoft.IdentityModel.Clients.ActiveDirectory*"}

    if($loadedAssemblies.Count -lt 2)
    {
        $global:loadedAssembliesStatus = Load-ActiveDirectoryAuthenticationLibrary
    }

    $clientId = "1950a258-227b-4e31-a9cf-717495945fc2"
    $redirectUri = "urn:ietf:wg:oauth:2.0:oob"
    $resourceClientId = "00000002-0000-0000-c000-000000000000"
    $authority = "https://login.windows.net/" + $tenant

    # Handle some edge cases
    if($env.ToLower() -eq "ppe")
    {
        $authority = "https://login.windows-ppe.net/" + $tenant
    }
    elseif($env.ToLower() -eq "china")
    {
        $authority = "https://login.chinacloudapi.cn/" + $tenant
    }

    # Set up our auth context and attempt to get a token
    $authContext = New-Object "Microsoft.IdentityModel.Clients.ActiveDirectory.AuthenticationContext" -ArgumentList $authority,$false
    $authResult = $authContext.AcquireToken($global:resourceManagerUri, $clientId, $redirectUri, [Microsoft.IdentityModel.Clients.ActiveDirectory.PromptBehavior]::Auto)

    # Return our auth result
    return $authResult
}

Function Connect-AzureAccount ($tenant = "common", $env="prod") 
{
    # Adapted from the excellent work here:
    #   http://www.dushyantgill.com/blog/2013/12/27/aadgraphpowershell-a-powershell-client-for-windows-azure-ad-graph-api/
    #   https://github.com/dushyantgill/AADGraphPowerShell/blob/master/AADGraph.psm1

    PROCESS {
        $global:aadAuthResult = $null
        $global:aadEnv = $env
        $global:subscriptionId = $null
        $global:resourceManagerUri = "https://management.core.windows.net/"
        $global:aadAuthResult = Get-AuthenticationResult -Tenant $tenant -Env $env

        # Set up our HTTP headers
        $header = $global:aadAuthResult.CreateAuthorizationHeader()
        $headers = @{"Authorization"=$header;"Content-Type"="application/json"}
        $uri = "https://management.azure.com/subscriptions?api-version=2015-01-01"

        # Query for list of subscriptions
        $subscriptions = (Invoke-RestMethod -Method Get -Uri $uri -Headers $headers -Verbose:$false).value | Select displayName,subscriptionId

        # Handle more than one Azure Subscription
        if ($subscriptions.Count -gt 1)
        {
            $caption = "Azure Subscriptions"
            $message = "Choose which Azure Subscription to draw:"
            $choiceList = @()
            $counter = 0
            foreach($subscription in $subscriptions)
            {
                $counter++
                $subscriptionName = $subscription.displayName
                $subscriptionId = $subscription.subscriptionId    
                $choice = New-Object System.Management.Automation.Host.ChoiceDescription "&$subscriptionName","$subscriptionId"
                $choiceList += $choice
            }
            $choices = [System.Management.Automation.Host.ChoiceDescription[]]($choiceList);
            $answer = $host.ui.PromptForChoice($caption,$message,$choices,0)

            Write-Host "You selected $($choiceList[$answer].Label.Substring(1))..."
            $global:subscriptionId = $choiceList[$answer].HelpMessage
        }
    }
}

Function Execute-ResourceManagerQuery($ApiAction, $QueryBase, $ApiVersion="2015-01-01", $QueryTail, $Data, [switch]$Verbose) 
{
    # Adapted from the excellent work here:
    #   http://www.dushyantgill.com/blog/2013/12/27/aadgraphpowershell-a-powershell-client-for-windows-azure-ad-graph-api/
    #   https://github.com/dushyantgill/AADGraphPowerShell/blob/master/AADGraph.psm1

    $response = $null

    # Check to make sure we have an existing auth result
    if($global:aadAuthResult -ne $null)
    {
        # Set up our HTTP headers
        $header = $global:aadAuthResult.CreateAuthorizationHeader()
        $headers = @{"Authorization"=$header;"Content-Type"="application/json"}

        # Set up the query string
        $uri = [string]::Format("https://management.azure.com/subscriptions/{0}/{1}?api-version={2}{3}",$global:subscriptionId, $QueryBase, $ApiVersion, $QueryTail)

        # Check to see if we're passing any data for a PUT action
        if($Data -ne $null)
        {
            $enc = New-Object "System.Text.ASCIIEncoding"
            $body = ConvertTo-Json -InputObject $Data -Depth 10
            $byteArray = $enc.GetBytes($body)
            $contentLength = $byteArray.Length
            $headers.Add("Content-Length",$contentLength)
        }

        # Dump out the planned HTTP action
        if($Verbose)
        {
          Write-Host "HTTP $ApiAction $uri`n" -ForegroundColor Cyan
        }
    
        # Dump the headers to be used in the HTTP action
        $headers.GetEnumerator() | % {
            if($Verbose)
            {
                Write-Host $_.Key: $_.Value -ForegroundColor Cyan
            }
        }

        # Dump data from a PUT action
        if($data -ne $null)
        {
            if($Verbose)
            {
                Write-Host "`n$body" -ForegroundColor Cyan
            }
        }

        # Execute the HTTP action
        $response = Invoke-WebRequest -Method $ApiAction -Uri $uri -Headers $headers -Body $body

        # Check the results of our action
        if($response.StatusCode -ge 200 -and $response.StatusCode -le 399)
        {
            if($Verbose)
            {
                Write-Host "`nQuery successfully executed." -ForegroundColor Cyan
            }
            if($response.Content -ne $null)
            {
                $json = ConvertFrom-Json $response.Content
                if($json -ne $null)
                {
                    $response = $json
                    if($json.value -ne $null)
                    {
                        $response = $json.value
                    }
                }
            }
        }
    }
    else
    {
        Write-Host "Not connected to an AAD tenant. First run Connect-AzureAccount." -ForegroundColor Yellow
    }
    return $response
}