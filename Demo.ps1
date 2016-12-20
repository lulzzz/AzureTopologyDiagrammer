param (

    # Turn off Verbose Logging
    $VerbosePreference = "SilentlyContinue",
    $Path = "D:\temp"
)



# Connect to Azure via ADAL
Connect-AzureAccount

# Fix up Visio Registry entries if needed
#Invoke-PatchOfficeC2RRegistry
#EndRegion

#Region Draw Resource Group Topology
Invoke-DrawAzureResourceGroups -Path $Path
#EndRegion


# Quit Visio
#$appInstance.Quit()