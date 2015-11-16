Function Extract-ZipFile($File, $Destination)
{
    # Create a new shell COM object
    $shell = New-Object -Com Shell.Application
    $zip = $shell.NameSpace($File)

    # Extract each item in the Zip file
    foreach($item in $zip.Items())
    {
        $shell.Namespace($Destination).CopyHere($item)
    }
}

Function Select-SubscriptionToDiagram
{
    # Get our Azure subscriptions
    $subscriptions = Get-AzureSubscription

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
            $subscriptionName = $subscription.SubscriptionName    
            $choice = New-Object System.Management.Automation.Host.ChoiceDescription "&$subscriptionName","$subscriptionName"
            $choiceList += $choice
        }
        $choices = [System.Management.Automation.Host.ChoiceDescription[]]($choiceList);
        $answer = $host.ui.PromptForChoice($caption,$message,$choices,0)

        $selectedSubscriptionName = $choiceList[$answer].HelpMessage

        Write-Host "You selected $selectedSubscriptionName..."
        Select-AzureSubscription -SubscriptionName "$selectedSubscriptionName"
    }
}

Function Patch-OfficeC2RRegistry
{
    # Check to see if we're running a ClickToRun version of Visio
    $usingC2R = Test-Path -Path "HKLM:SOFTWARE\Microsoft\Office\ClickToRun"
    if ($usingC2R)
    {
        [bool]$testKey1 = Test-Path -Path "HKLM:\SOFTWARE\Classes\CLSID\{00021A20-0000-0000-C000-000000000046}"
        [bool]$testKey2 = Test-Path -Path "HKLM:\SOFTWARE\Classes\Wow6432Node\CLSID\{00021A20-0000-0000-C000-000000000046}"
        [bool]$testKey3 = Test-Path -Path "HKLM:\SOFTWARE\Classes\Interface\{000D0700-0000-0000-C000-000000000046}"
        [bool]$testResults = ($testKey1 -and $testKey2 -and $testKey3)

        if(!$testResults)
        {
            Write-Host -ForegroundColor Yellow "You're using Office Click2Run, so we need to fix some registry keys..."
            $registryKeyMods = '
            Copy-Item -Path "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Classes\CLSID\{00021A20-0000-0000-C000-000000000046}" -Destination "HKLM:\SOFTWARE\Classes\CLSID\{00021A20-0000-0000-C000-000000000046}" -Recurse -Force
            Copy-Item -Path "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Classes\Wow6432Node\CLSID\{00021A20-0000-0000-C000-000000000046}" -Destination "HKLM:\SOFTWARE\Classes\Wow6432Node\CLSID\{00021A20-0000-0000-C000-000000000046}" -Recurse -Force
            Copy-Item -Path "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Classes\Interface\{000D0700-0000-0000-C000-000000000046}" -Destination "HKLM:\SOFTWARE\Classes\Interface\{000D0700-0000-0000-C000-000000000046}" -Recurse -Force
            '
            $encodedCommand = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($registryKeyMods))
            Start-Process -FilePath powershell.exe -Verb runas -ArgumentList "-encodedCommand $encodedCommand"
        }
    }
}

Function Get-VisioCloudStencils
{
    # Downloads Cloud and Enterprise Visio Stencils and extracts them to a subfolder
    #   Download Center: http://www.microsoft.com/en-us/download/details.aspx?id=41937
    #   Direct: https://download.microsoft.com/download/1/7/1/171DA19A-5477-4F50-B354-4ABAF28502A6/Microsoft_CloudnEnterprise_Symbols_v2.3_Public.zip

    $cloudStencilUrl = "https://download.microsoft.com/download/1/7/1/171DA19A-5477-4F50-B354-4ABAF28502A6/Microsoft_CloudnEnterprise_Symbols_v2.3_Public.zip"
    $cloudStencilPath = "$pwd\CnE_VisioStencils"
    $cloudStencilZipPath = "$pwd\MicrosoftCloudSymbols.zip"

    # Check for extracted Cloud & Enterprise stencils
    if (!(Test-Path -Path $cloudStencilPath))
    {
        # Create the directory for the Cloud & Enterprise stencils
        New-Item -ItemType Directory -Path $cloudStencilPath

        # Check to see if we've already downloaded the stencils
        if (!(Test-Path -Path $cloudStencilZipPath))
        {
            # If no downloaded stencils, download using BITS
            Start-BitsTransfer -Source $cloudStencilUrl -Destination "$pwd\MicrosoftCloudSymbols.zip"
        }

        # Ensure downloaded zip isn't blocked, then extract
        Unblock-File -Path $cloudStencilZipPath
        Extract-ZipFile -File $cloudStencilZipPath -Destination $cloudStencilPath
    }
}

Function Draw-AzureResourceGroups($VisioPage)
{
    # Name our current page
    $VisioPage.Name = "All Resource Groups"

    # Switch to Azure Resource Manager to get our Resource Groups
    Switch-AzureMode AzureResourceManager -ErrorAction SilentlyContinue
    $resourceGroups = Get-AzureResourceGroup
    $geoLocations = $resourceGroups | Select-Object -Unique Location

    # Enable Diagram Services
    # Reference: https://msdn.microsoft.com/en-us/library/office/ff765437.aspx
    $visioDocument.DiagramServicesEnabled = 8

    $geoCounter = 1
    foreach($geoLocation in $geoLocations)
    {
        # Add Region objects    
        $cloudShape = $VisioPage.Drop($cloudMaster, $geoCounter, 1)
        $cloudShape.Text = $geoLocation.Location

        # Get the actual cloud ID for fill
        $cloudShapeID = $cloudShape.ID + 1
        $cloudShape.Shapes.ItemFromID($cloudShapeID).Cells("FillForegnd").Formula = "THEMEGUARD(RGB(0,120,215))"

        $geoCounter++
    }

    $resourceGroupCounter = 0
    foreach($resourceGroup in $resourceGroups)
    {
        # Get relevant Region object
        $regionShape = $VisioPage.Shapes | Where-Object {$_.Text -eq $resourceGroup.Location}
        $regionShapeX = $regionShape.CellsU("PinX").ResultIU
        $regionShapeY = $regionShape.CellsU("PinY").ResultIU

        # Add Resource Group object    
        $resourceGroupShape = $VisioPage.Drop($resourceGroupMaster, $regionShapeX, $regionShapeY)
        $resourceGroupShape.Text = $resourceGroup.ResourceGroupName
        $resourceGroupShape.CellsSRC(3,0,1).FormulaU = "THEMEGUARD(RGB(0,120,215))"
        $resourceGroupShape.Cells("Width").Formula = "MIN(TEXTWIDTH($($resourceGroupShape.Name)!theText,2),2)"
        $resourceGroupShape.Cells("Height").Formula = "TEXTHEIGHT($($resourceGroupShape.Name)!theText,$($resourceGroupShape.Cells("Width").ResultIU))"
        $resourceGroupShape.Cells("LineColor").Formula = "THEMEGUARD(RGB(0,120,215))"
    
        # Connect Resource Group to Region
        $connector = $VisioPage.Drop($VisioPage.Application.ConnectorToolDataObject,0,0)
        $connector.CellsU("LineColor").Formula = "THEMEGUARD(RGB(150,150,150))"
        $startX = $connector.CellsU("BeginX").GlueTo($resourceGroupShape.CellsU("PinX"))
        $startY = $connector.CellsU("BeginY").GlueTo($resourceGroupShape.CellsU("PinY"))
        $endX = $connector.CellsU("EndX").GlueTo($regionShape.CellsU("PinX"))
        $endY = $connector.CellsU("EndY").GlueTo($regionShape.CellsU("PinY"))
    }
    
    # Configure Layout and Routing Styles
    #   RouteStyle: https://msdn.microsoft.com/en-us/library/office/ff765968.aspx
    #   LineRoute: https://msdn.microsoft.com/en-us/library/office/ff766029.aspx
    $VisioPage.PageSheet.CellsU("RouteStyle").ResultIU = 5 
    $VisioPage.PageSheet.CellsU("LineRouteExt").ResultIU = 2 

    # Auto layout and size before adding callouts for Resource Group tags
    $VisioPage.Layout()
    $VisioPage.ResizeToFitContents()

    $resourceGroupCounter = 0
    foreach($resourceGroup in $resourceGroups)
    {
        # Get relevant Resource Group object    
        $resourceGroupShape = $VisioPage.Shapes | Where-Object {$_.Text -eq $resourceGroup.ResourceGroupName}
        $resourceGroupShapeX = $resourceGroupShape.CellsU("PinX").ResultIU
        $resourceGroupShapeY = $resourceGroupShape.CellsU("PinY").ResultIU

        # Get our resource group tags
        $resourceGroupTags = $resourceGroup.Tags

        # If any tags, add callout with tag details
        if ($resourceGroupTags.Count -gt 0)
        {
            $resourceGroupTagShape = $VisioPage.DropCallout($calloutMaster, $resourceGroupShape)
            $resourceGroupTagShapeText = ""

            foreach($resourceGroupTag in $resourceGroupTags.GetEnumerator())
            {
                $resourceGroupTagShapeText += "`n$($resourceGroupTag.Name) : $($resourceGroupTag.Value)"
            }

            $resourceGroupTagShape.Text = $resourceGroupTagShapeText
            $resourceGroupTagShape.CellsSRC(3,0,1).FormulaU = "THEMEGUARD(RGB(0,120,215))"
        }  
    }
}