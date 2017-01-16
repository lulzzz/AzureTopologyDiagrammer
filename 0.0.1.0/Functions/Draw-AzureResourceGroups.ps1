Function Invoke-DrawAzureResourceGroups {
    param (
        $Path
    )

    BEGIN {
        # Load DLL via reflection
        $loadVisioDll = [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.Office.Interop.Visio")

        # Create a Visio Application Object
        $appObject = New-Object Microsoft.Office.Interop.Visio.ApplicationClass
        $appInstance = $appObject.Application
        $appDocuments = $appInstance.Documents

        # Add a new document from the Detailed Network Diagram template
        $visioDocument = $appDocuments.Add("DTLNME_U.VSTX")

        # Set file path for the Save action
        $azureVisioPath = Join-Path $Path "AzureTopology.vsdx"

        # Setup our stencils
        $azureVisioCloudStencilPath = Get-ChildItem -Path $Script:azureVisioStencilFolder -Filter "CnE_Cloud*"
        $azureVisioCloudStencil = $appDocuments.Add($($azureVisioCloudStencilPath.FullName))
        $networkLocationsStencil = $appDocuments | Where-Object {$_.Title -eq "Network Locations"}

        # Add our Callouts and Containers Stencils
        # Reference: https://msdn.microsoft.com/en-us/library/office/ff765723.aspx
        $builtInCalloutsStencilPath = $appInstance.GetBuiltInStencilFile(3,0)
        $builtInCalloutsStencil = $appDocuments.Add($builtInCalloutsStencilPath)
        $builtInContainersStencilPath = $appInstance.GetBuiltInStencilFile(2,0)
        $builtInContainersStencil = $appDocuments.Add($builtInContainersStencilPath)

        # Setup our Masters
        $cloudMaster = $networkLocationsStencil.Masters.Item("Cloud")
        $resourceGroupMaster = $azureVisioCloudStencil.Masters.Item("Affinity group")
        $virtualNetworkMaster = $azureVisioCloudStencil.Masters.Item("Virtual Network")
        $virtualNetworkBoxMaster = $azureVisioCloudStencil.Masters.Item("Virtual Network Box")
        $calloutMaster = $builtInCalloutsStencil.Masters.Item("Orthogonal")
        $containerMaster = $builtInContainersStencil.Masters.Item("Plain")

        # Get our Visio pages
        $visioPages = $visioDocument.Pages
        $resourceGroupPage = $visioPages.Item(1)
        $VisioPage= $resourceGroupPage
    }
    PROCESS {
        # Name our current page
        $VisioPage.Name = "All Resource Groups"

        # Get the unique Azure deployment locations
        $geoLocations = Get-ResourceGroups | Select-Object -Unique Location

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
    END {
        # Save the Visio
        $visioDocument.SaveAs($azureVisioPath)
    }
}


Export-ModuleMember Invoke-DrawAzureResourceGroups