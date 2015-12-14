# ===================================================================================
# Name:    Azure Topology Diagrammer
# Description: Functions for Visio diagramming
# ===================================================================================

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

Function Patch-OfficeC2RRegistry
{
    # Check to see if we're running a ClickToRun version of Visio
    $usingC2R = Test-Path -Path "HKLM:SOFTWARE\Microsoft\Office\ClickToRun"
    if ($usingC2R)
    {
        # Check to make sure registry entries are present
        [bool]$testKey1 = Test-Path -Path "HKLM:\SOFTWARE\Classes\CLSID\{00021A20-0000-0000-C000-000000000046}"
        [bool]$testKey2 = Test-Path -Path "HKLM:\SOFTWARE\Classes\Wow6432Node\CLSID\{00021A20-0000-0000-C000-000000000046}"
        [bool]$testKey3 = Test-Path -Path "HKLM:\SOFTWARE\Classes\Interface\{000D0700-0000-0000-C000-000000000046}"
        [bool]$testResults = ($testKey1 -and $testKey2 -and $testKey3)

        # If missing registry entries, patch
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

    # Get our Resource Groups
    $resourceGroups = Get-ResourceGroups

    # Get the unique Azure deployment locations
    $geoLocations = $resourceGroups | Select-Object -Unique Location

    # Enable Diagram Services
    $visioDocument.DiagramServicesEnabled = $visServiceVersion150

    $geoCounter = 1
    foreach($geoLocation in $geoLocations)
    {
        # Add Region objects    
        $cloudShape = $VisioPage.Drop($cloudMaster, $geoCounter, 1)
        $cloudShape.Text = $geoLocation.Location

        # Get the actual cloud ID for fill
        $cloudShapeID = $cloudShape.ID + 1
        $cloudShape.Shapes.ItemFromID($cloudShapeID).Cells("FillForegnd").Formula = "THEMEGUARD(RGB($rgbAzure))"

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
        $resourceGroupShape.Text = $resourceGroup.Name
        $resourceGroupShape.CellsSRC($visSectionCharacter,$visRowCharacter,$visCharacterColor).FormulaU = "THEMEGUARD(RGB($rgbAzure))"
        $resourceGroupShape.Cells("Width").Formula = "MIN(TEXTWIDTH($($resourceGroupShape.Name)!theText,2),2)"
        $resourceGroupShape.Cells("Height").Formula = "TEXTHEIGHT($($resourceGroupShape.Name)!theText,$($resourceGroupShape.Cells("Width").ResultIU))"
        $resourceGroupShape.Cells("LineColor").Formula = "THEMEGUARD(RGB($rgbAzure))"
    
        # Connect Resource Group to Region
        $connector = $VisioPage.Drop($VisioPage.Application.ConnectorToolDataObject,0,0)
        $connector.CellsU("LineColor").Formula = "THEMEGUARD(RGB($rgbGeneral))"
        $startX = $connector.CellsU("BeginX").GlueTo($resourceGroupShape.CellsU("PinX"))
        $startY = $connector.CellsU("BeginY").GlueTo($resourceGroupShape.CellsU("PinY"))
        $endX = $connector.CellsU("EndX").GlueTo($regionShape.CellsU("PinX"))
        $endY = $connector.CellsU("EndY").GlueTo($regionShape.CellsU("PinY"))
    }
    
    # Configure Layout and Routing Styles
    #   RouteStyle: https://msdn.microsoft.com/en-us/library/office/ff765968.aspx    
    $VisioPage.PageSheet.CellsSRC($visSectionObject,$visRowRulerGrid,$visXRulerOrigin).FormulaU = "0 in"
    $VisioPage.PageSheet.CellsSRC($visSectionObject,$visRowRulerGrid,$visYRulerOrigin).FormulaU = "0 in"
    $VisioPage.PageSheet.CellsSRC($visSectionObject,$visRowRulerGrid,$visXGridOrigin).FormulaU = "0 in"
    $VisioPage.PageSheet.CellsSRC($visSectionObject,$visRowRulerGrid,$visYGridOrigin).FormulaU = "0 in"
    $VisioPage.PageSheet.CellsSRC($visSectionObject,$visRowPageLayout,$visPLOPlaceStyle).FormulaForceU = "1"
    $VisioPage.PageSheet.CellsSRC($visSectionObject,$visRowPageLayout,$visPLORouteStyle).FormulaForceU = "5"

    # Configure smooth line style
    #   LineRoute: https://msdn.microsoft.com/en-us/library/office/ff766029.aspx
    $VisioPage.PageSheet.CellsU("LineRouteExt").ResultIU = 2

    # Auto layout and resize page before adding callouts for Resource Group tags
    $VisioPage.Layout()
    $VisioPage.ResizeToFitContents()

    $resourceGroupCounter = 0
    foreach($resourceGroup in $resourceGroups)
    {
        # Get the resource group tags
        $resourceGroupTagsAsString = $resourceGroup.Tags

        if($resourceGroupTagsAsString -ne $null)
        {
            # Convert the PSCustomObject to a hashtable
            $resourceGroupTags = @{}
            $resourceGroupTagsAsString.PSObject.Properties | Foreach { $resourceGroupTags[$_.Name] = $_.Value }

            # If any tags, get the resource group shape and add callout with tag details
            if ($resourceGroupTags.Count -gt 0)
            {
                # Get relevant Resource Group object    
                $resourceGroupShape = $VisioPage.Shapes | Where-Object {$_.Text -eq $resourceGroup.Name}
                $resourceGroupShapeX = $resourceGroupShape.CellsU("PinX").ResultIU
                $resourceGroupShapeY = $resourceGroupShape.CellsU("PinY").ResultIU

                $resourceGroupTagShape = $VisioPage.DropCallout($calloutMaster, $resourceGroupShape)
                $resourceGroupTagShapeText = "Tags"

                foreach($resourceGroupTag in $resourceGroupTags.GetEnumerator())
                {
                    $resourceGroupTagShapeText += "`n    $($resourceGroupTag.Name) : $($resourceGroupTag.Value)"
                }

                $resourceGroupTagShape.Text = $resourceGroupTagShapeText
                $resourceGroupTagShape.CellsSRC($visSectionCharacter,$visRowCharacter,$visCharacterColor).FormulaU = "THEMEGUARD(RGB($rgbGeneral))"
                $resourceGroupTagShape.CellsSRC($visSectionParagraph,$visRowParagraph,$visHorzAlign).FormulaU = 0
                $resourceGroupTagShape.CellsSRC($visSectionObject,$visRowLine,$visLineColor).FormulaU = "THEMEGUARD(RGB($rgbGeneral))"
                $resourceGroupTagShape.Cells("Width").Formula = "MIN(TEXTWIDTH($($resourceGroupTagShape.Name)!theText,2),2)"
            }  
        }
    }
}


Function Draw-AzureNetworkDetails($VisioPage)
{
    # Name our current page
    $VisioPage.Name = "Network Details"

    # Retrieve both v1 and v2 networks
    $networks = Get-Networks
    $classicNetworks = Get-ClassicNetworks

    # Get our network locations from v1 and v2 networks
    $networkLocations = $networks | Select-Object -Unique Location
    $networkLocations += $classicNetworks | Select-Object -Unique Location

    # Make sure we only have unique locations
    $geoLocations = $networkLocations | Select-Object -Unique Location

    # Enable Diagram Services
    $visioDocument.DiagramServicesEnabled = $visServiceVersion150

    # Draw the geo regions
    $geoCounter = 1
    foreach($geoLocation in $geoLocations)
    {
        # Add Region objects    
        $cloudShape = $VisioPage.Drop($cloudMaster, $geoCounter, 7)
        $cloudShape.Text = $geoLocation.Location

        # Get the actual cloud ID for fill
        $cloudShapeID = $cloudShape.ID + 1
        $cloudShape.Shapes.ItemFromID($cloudShapeID).Cells("FillForegnd").Formula = "THEMEGUARD(RGB($rgbAzure))"

        # Create a transparent container to hold related networks (due to formatting issues)
        $containerShape = $VisioPage.Drop($containerMaster, $geoCounter, 1)
        $containerShape.Text = "$($geoLocation.Location)-Networks" 
        #$containerShape.CellsU("User.msvStructureType").FormulaU = '"List"'
        #$containerShape.CellsU("LinePattern").ResultIU = 0
        #$containerShapeID = $containerShape.ID + 1
        $containerShapeTextID = $containerShape.ID + 3
        #$containerShape.Shapes.ItemFromID($containerShapeID).CellsU("LineColor").Formula = "THEMEGUARD(RGB($rgbAzure))"
        $containerShape.Shapes.ItemFromID($containerShapeTextID).CellsSRC($visSectionCharacter,$visRowCharacter,$visCharacterColor).FormulaU = "THEMEGUARD(RGB(255,255,255))"    
        # Visio Unit Codes: https://msdn.microsoft.com/en-us/library/office/ff769148.aspx
        # Visio SetListSpacing: https://msdn.microsoft.com/en-us/library/office/ff765721.aspx
        #$containerShape.ContainerProperties.SetListSpacing($visInches,.0125)
        #$containerShape.ContainerProperties.ListDirection = $visListDirTopToBottom
        $containerShape.Cells("Width").Formula = "1.5"        
        $containerShape.Cells("Height").Formula = "1"
        $containerShape.ContainerProperties.ResizeAsNeeded = 2
        
        $geoCounter = $geoCounter + 2.5
    }
        
    # Configure Layout and Routing Styles
    #   RouteStyle: https://msdn.microsoft.com/en-us/library/office/ff765968.aspx    
    $VisioPage.PageSheet.CellsSRC($visSectionObject,$visRowRulerGrid,$visXRulerOrigin).FormulaU = "0 in"
    $VisioPage.PageSheet.CellsSRC($visSectionObject,$visRowRulerGrid,$visYRulerOrigin).FormulaU = "0 in"
    $VisioPage.PageSheet.CellsSRC($visSectionObject,$visRowRulerGrid,$visXGridOrigin).FormulaU = "0 in"
    $VisioPage.PageSheet.CellsSRC($visSectionObject,$visRowRulerGrid,$visYGridOrigin).FormulaU = "0 in"
    $VisioPage.PageSheet.CellsSRC($visSectionObject,$visRowPageLayout,$visPLOPlaceStyle).FormulaForceU = "1"
    $VisioPage.PageSheet.CellsSRC($visSectionObject,$visRowPageLayout,$visPLORouteStyle).FormulaForceU = "5"

    # Configure smooth line style
    #   LineRoute: https://msdn.microsoft.com/en-us/library/office/ff766029.aspx
    $VisioPage.PageSheet.CellsU("LineRouteExt").ResultIU = 2

    # Auto layout and resize page before adding networks and subnets to hidden containers
    #$VisioPage.Layout()
    
    $lastContainers = @()

    # Draw the virtual networks
    foreach($network in $networks)
    {
        $networkLocation = $network.location
        $networkName = $network.Name
        
        # Get relevant region
        $regionShape = $VisioPage.Shapes | Where-Object {$_.Text -eq $networkLocation}

        # Get relevant hidden container
        $hiddenContainer = $VisioPage.Shapes | Where-Object {$_.Text -eq "$networkLocation-Networks"}
        $hiddenContainerX = $hiddenContainer.CellsU("PinX").ResultIU
        $hiddenContainerY = $hiddenContainer.CellsU("PinY").ResultIU

        # Look for existing networks added for this geographic location
        $lastGeoMatch = $lastContainers | Where-Object {$_.GeoLocation -eq $networkLocation} | Select-Object -Last 1

        # Initialize variables
        $containerShape = $null
        $locationY = $null

        # If an existing network was found on the page...
        if ($lastGeoMatch)
        {
            # Find the last container Y location and height
            [int]$origLocation = $lastGeoMatch.LocationY
            [int]$origHeight = $lastGeoMatch.Height

            # Compute the new center location for our container and add it to the page
            $locationY =  $origLocation - ($origHeight / 2) - 1
            $containerShape = $VisioPage.Drop($containerMaster, $hiddenContainerX, $locationY)
        }
        else
        {
            # Create a container representing a Virtual Network
            $locationY = $hiddenContainerY
            $containerShape = $VisioPage.Drop($containerMaster, $hiddenContainerX, $locationY)
        }

        $containerShape.Text = $networkName    
        $containerShape.CellsU("User.msvStructureType").FormulaU = '"List"'
        $containerShape.CellsU("LinePattern").ResultIU = 9
        $containerShapeID = $containerShape.ID + 1
        $containerShapeTextID = $containerShape.ID + 3
        $containerShape.Shapes.ItemFromID($containerShapeID).CellsU("LineColor").Formula = "THEMEGUARD(RGB($rgbAzure))"
        $containerShape.Shapes.ItemFromID($containerShapeTextID).CellsSRC($visSectionCharacter,$visRowCharacter,$visCharacterColor).FormulaU = "THEMEGUARD(RGB($rgbAzure))"
        $containerShape.ContainerProperties.SetListSpacing($visInches,.0125)
        $containerShape.ContainerProperties.ListDirection = $visListDirTopToBottom
        $containerShape.ContainerProperties.ResizeAsNeeded = 1
        $containerShape.Cells("User.msvSDContainerMargin") = .025
                                
        # Insert the virtual network container into the hidden container
        $hiddenContainer.ContainerProperties.AddMember($containerShape, 1)

        # Retrieve the network subnets
        $networkSubnets = Get-NetworkSubnets -Network $network

        # Draw the network subnets
        $counter = 1
        foreach($networkSubnet in $networkSubnets.GetEnumerator())
        {
            $subnetShape = $VisioPage.Drop($virtualNetworkBoxMaster, $hiddenContainerX, $locationY)        
            $subnetShape.Text = "$($networkSubnet.Name)`n$($networkSubnet.Value)"
            $subnetShape.CellsSRC($visSectionCharacter,$visRowCharacter,$visCharacterColor).FormulaU = "THEMEGUARD(RGB($rgbAzure))"
            $subnetShape.Cells("Width").Formula = "1.9" #"MIN(TEXTWIDTH($($subnetShape.Name)!theText,1.8),1.8)"
            $subnetShape.Cells("Height").Formula = "TEXTHEIGHT($($subnetShape.Name)!theText,$($subnetShape.Cells("Width").ResultIU))"
            
            # After the subnets are added, shift the center point down the page by half the height
            [int]$containerHeight = $containerShape.Cells("Height").ResultIU
            $containerShape.CellsU("PinY").ResultIU = $locationY - ($containerHeight/2)
            $subnetShape.CellsU("PinY").ResultIU = $locationY - ($containerHeight/2)

            $containerShape.ContainerProperties.InsertListMember($subnetShape, $counter)

            $counter++
        }



        # Connect hidden container to Region
        $connector = $VisioPage.Drop($VisioPage.Application.ConnectorToolDataObject,0,0)
        $connector.CellsU("LineColor").Formula = "THEMEGUARD(RGB($rgbGeneral))"
        $startX = $connector.CellsU("BeginX").GlueTo($containerShape.CellsU("PinX"))
        $startY = $connector.CellsU("BeginY").GlueTo($containerShape.CellsU("PinY"))
        $endX = $connector.CellsU("EndX").GlueTo($regionShape.CellsU("PinX"))
        $endY = $connector.CellsU("EndY").GlueTo($regionShape.CellsU("PinY"))

        # Create an object definition for our network containers
        $lastContainer = New-Object -TypeName PSObject
        $lastContainer | Add-Member -MemberType NoteProperty -Name LocationY -Value $containerShape.CellsU("PinY").ResultIU
        $lastContainer | Add-Member -MemberType NoteProperty -Name GeoLocation -Value $networkLocation
        $lastContainer | Add-Member -MemberType NoteProperty -Name Height -Value $containerHeight

        Write-Host "Saved $($lastContainer.Height) height and PinY $($lastContainer.LocationY)"
        
        # Add the object to an array
        $lastContainers += $lastContainer

        # Flush our container object
        $lastContainer = $null
    }

    <#

    # Draw the classic virtual networks
    foreach($classicNetwork in $classicNetworks)
    {
        $networkLocation = $classicNetwork.location
        $networkName = $classicNetwork.Name

         # Get relevant region
        $regionShape = $VisioPage.Shapes | Where-Object {$_.Text -eq $networkLocation}

        # Get relevant hidden container
        $hiddenContainer = $VisioPage.Shapes | Where-Object {$_.Text -eq "$networkLocation-Networks"}
        $hiddenContainerX = $hiddenContainer.CellsU("PinX").ResultIU
        $hiddenContainerY = $hiddenContainer.CellsU("PinY").ResultIU

        # Create a container representing a Virtual Network
        $containerShape = $VisioPage.Drop($containerMaster, $hiddenContainerX, $hiddenContainerY)
        $containerShape.Text = $networkName    
        $containerShape.CellsU("User.msvStructureType").FormulaU = '"List"'
        $containerShape.CellsU("LinePattern").ResultIU = 9
        $containerShapeID = $containerShape.ID + 1
        $containerShapeTextID = $containerShape.ID + 3
        $containerShape.Shapes.ItemFromID($containerShapeID).CellsU("LineColor").Formula = "THEMEGUARD(RGB($rgbAzure))"
        $containerShape.Shapes.ItemFromID($containerShapeTextID).CellsSRC($visSectionCharacter,$visRowCharacter,$visCharacterColor).FormulaU = "THEMEGUARD(RGB($rgbAzure))"
        $containerShape.ContainerProperties.SetListSpacing($visInches,.0125)
        $containerShape.ContainerProperties.ListDirection = $visListDirTopToBottom
        $containerShape.ContainerProperties.ResizeAsNeeded = 1
        $containerShape.Cells("User.msvSDContainerMargin") = .025
                                
        # Insert the virtual network container into the hidden container
        $hiddenContainer.ContainerProperties.AddMember($containerShape, 1)

        # Retrieve the network subnets
        $networkSubnets = Get-ClassicNetworkSubnets -ClassicNetwork $classicNetwork

        # Draw the network subnets
        $counter = 1
        foreach($networkSubnet in $networkSubnets.GetEnumerator())
        {
            $subnetShape = $VisioPage.Drop($virtualNetworkBoxMaster, $hiddenContainerX, $hiddenContainerY)        
            $subnetShape.Text = "$($networkSubnet.Name)`n$($networkSubnet.Value)"
            $subnetShape.CellsSRC($visSectionCharacter,$visRowCharacter,$visCharacterColor).FormulaU = "THEMEGUARD(RGB($rgbAzure))"
            $subnetShape.Cells("Width").Formula = "1.9" #"MIN(TEXTWIDTH($($subnetShape.Name)!theText,1.8),1.8)"
            $subnetShape.Cells("Height").Formula = "TEXTHEIGHT($($subnetShape.Name)!theText,$($subnetShape.Cells("Width").ResultIU))"

            $containerShape.ContainerProperties.InsertListMember($subnetShape, $counter)
            $counter++
        }

        # Connect hidden container to Region
        $connector = $VisioPage.Drop($VisioPage.Application.ConnectorToolDataObject,0,0)
        $connector.CellsU("LineColor").Formula = "THEMEGUARD(RGB($rgbGeneral))"
        $startX = $connector.CellsU("BeginX").GlueTo($containerShape.CellsU("PinX"))
        $startY = $connector.CellsU("BeginY").GlueTo($containerShape.CellsU("PinY"))
        $endX = $connector.CellsU("EndX").GlueTo($regionShape.CellsU("PinX"))
        $endY = $connector.CellsU("EndY").GlueTo($regionShape.CellsU("PinY"))
    }

    #>

    # Resize to fit contents
    $VisioPage.ResizeToFitContents()
}