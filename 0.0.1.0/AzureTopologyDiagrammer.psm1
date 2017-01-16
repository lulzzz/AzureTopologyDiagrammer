
# Loading libraries
$LibrariesPath = "$PSScriptRoot\lib"
Get-ChildItem $LibrariesPath -filter "*.dll" | ForEach-Object {
    (join-path $LibrariesPath "$($PSItem.BaseName).ps1")
}

# dot sourcing all functions
$FunctionsPath = "$PSScriptRoot\Functions"
Get-ChildItem $FunctionsPath | ForEach-Object {
    . (join-path $FunctionsPath "$($PSItem.BaseName).ps1")
}

# Initializing some vars
[string]$Script:CurrentSubscriptionId = $null
[string]$Script:AuthToken = $null
[string]$Script:RefreshToken = $null
[Nullable[System.DateTimeOffset]]$Script:TokenExpirationUtc = $null
$Script:clientId = "1950a258-227b-4e31-a9cf-717495945fc2"
$Script:LoginUrl = "https://login.microsoftonline.com/common/oauth2/authorize"
$Script:redirectUri = "urn:ietf:wg:oauth:2.0:oob"
$Script:ResourceUrl = "https://management.core.windows.net/"

<#
# Removing some nasty verbose output
$Script:PsDefaultParameterValues.Add("Invoke-RestMethod:Verbose",$False)
$Script:PsDefaultParameterValues.Add("Invoke-WebRequest:Verbose",$False)
#>

# Initializing Visio vars
# RGB Codes
$Script:rgbAzure = "0,120,215"
$Script:rgbGeneral = "150,150,150"
$Script:rgbOffice365 = "220,60,0"
$Script:rgbOnPrem = "0,24,143"

# visSectionIndices:  https://msdn.microsoft.com/EN-US/library/office/ff765983.aspx    
[int]$Script:visSectionObject = 1
[int]$Script:visSectionCharacter = 3
[int]$Script:visSectionParagraph = 4

# visRowIndices:      https://msdn.microsoft.com/EN-US/library/office/ff765539.aspx
[int]$Script:visRowRulerGrid = 18
[int]$Script:visRowPageLayout = 24
[int]$Script:visRowCharacter = 0
[int]$Script:visRowLine = 2
[int]$Script:visRowParagraph = 0

# visCellIndices:     https://msdn.microsoft.com/EN-US/library/office/ff767991.aspx
[int]$Script:visCharacterColor = 1
[int]$Script:visCharacterDblUnderline = 8
[int]$Script:visHorzAlign = 6
[int]$Script:visLineColor = 1
[int]$Script:visPLOPlaceStyle = 8
[int]$Script:visPLORouteStyle = 9
[int]$Script:visXRulerOrigin = 4
[int]$Script:visXGridOrigin = 10
[int]$Script:visYRulerOrigin = 5
[int]$Script:visYGridOrigin = 11

# visDiagramServices: https://msdn.microsoft.com/en-us/library/office/ff768414(v=office.15).aspx
[int]$Script:visServiceVersion150 = 8

# Downloading Visio stencils if needed
$cloudStencilUrl = "https://download.microsoft.com/download/1/7/1/171DA19A-5477-4F50-B354-4ABAF28502A6/Microsoft_CloudnEnterprise_Symbols_v2.3_Public.zip"
$cloudStencilPath = Join-Path $PSScriptRoot "CnE_VisioStencils"
$cloudStencilZipPath = Join-Path $PSScriptRoot "MicrosoftCloudSymbols.zip"
$Script:azureVisioStencilFolder = Join-Path (Join-Path $PSScriptRoot "CnE_VisioStencils") "Visio"

if (!(Test-Path -Path "$cloudStencilPath")) {
    # Create the directory for the Cloud & Enterprise stencils
    New-Item -ItemType Directory -Path $cloudStencilPath | Out-Null

    # Check to see if we've already downloaded the stencils
    if (!(Test-Path -Path $cloudStencilZipPath)) {
        # If no downloaded stencils, download using BITS
        Start-BitsTransfer -Source $cloudStencilUrl -Destination $cloudStencilZipPath
    }

    # Ensure downloaded zip isn't blocked, then extract
    Unblock-File -Path $cloudStencilZipPath
    Expand-Archive -Path $cloudStencilZipPath -DestinationPath $cloudStencilPath -Force
}