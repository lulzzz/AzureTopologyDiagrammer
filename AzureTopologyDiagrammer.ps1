# ===================================================================================
# Name:    Azure Topology Diagrammer
# Version: 1.0
# Author:  Wes Kroesbergen
# Web:     http://www.kroesbergens.com
# ===================================================================================

# Turn off Verbose Logging
$VerbosePreference = "SilentlyContinue"

#Region Setup Paths & Environment
$Host.UI.RawUI.WindowTitle = " -- Azure Topology Diagrammer -- by Wes Kroesbergen --"

# Dot source our functions and enums
. "$pwd\AzureTopologyDiagrammerAzureHelpers.ps1"
. "$pwd\AzureTopologyDiagrammerEnumsAndVars.ps1"
. "$pwd\AzureTopologyDiagrammerFunctions.ps1"

# Connect to Azure via ADAL
Connect-AzureAccount

# Fix up Visio Registry entries if needed
Patch-OfficeC2RRegistry

# Make sure we have the Visio Cloud and Enterprise Stencils available
Get-VisioCloudStencils
#EndRegion

#Region Initialize Visio
# Load DLL via reflection
$loadVisioDll = [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.Office.Interop.Visio")

# Create a Visio Application Object
$appObject = New-Object Microsoft.Office.Interop.Visio.ApplicationClass
$appInstance = $appObject.Application
$appDocuments = $appInstance.Documents

# Add a new document from the Detailed Network Diagram template
$visioDocument = $appDocuments.Add("DTLNME_U.VSTX")

# Add the Azure Stencil set
$azureVisioStencilFolder = "$pwd\CnE_VisioStencils\Visio"

# Set file path for the Save action
$azureVisioPath = "$pwd\AzureTopology.vsdx"

# Setup our stencils
$azureVisioCloudStencilPath = Get-ChildItem -Path $azureVisioStencilFolder -Filter "CnE_Cloud*"
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
#EndRegion

#Region Draw Resource Group Topology
$resourceGroupPage = $visioPages.Item(1)
Draw-AzureResourceGroups -VisioPage $resourceGroupPage
#EndRegion

#Region Draw Virtual Network Topology
$allNetworksPage = $visioPages.Add()
Draw-AzureNetworkDetails -VisioPage $allNetworksPage
#EndRegion

# Save our changes
$visioDocument.SaveAs($azureVisioPath)

# Quit Visio
#$appInstance.Quit()