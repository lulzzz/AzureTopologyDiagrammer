Function Invoke-PatchOfficeC2RRegistry {
    # Check to see if we're running a ClickToRun version of Visio
    $usingC2R = Test-Path -Path "HKLM:SOFTWARE\Microsoft\Office\ClickToRun"
    if ($usingC2R)
    {
        # Check to make sure registry entries are present
        $testKey1 = Test-Path -Path "HKLM:\SOFTWARE\Classes\CLSID\{00021A20-0000-0000-C000-000000000046}"
        $testKey2 = Test-Path -Path "HKLM:\SOFTWARE\Classes\Wow6432Node\CLSID\{00021A20-0000-0000-C000-000000000046}"
        $testKey3 = Test-Path -Path "HKLM:\SOFTWARE\Classes\Interface\{000D0700-0000-0000-C000-000000000046}"
        $testResults = ($testKey1 -and $testKey2 -and $testKey3)

        # If missing registry entries, patch
        # Copy-Item -Path "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Classes\Interface\{000D0700-0000-0000-C000-000000000046}" -Destination "HKLM:\SOFTWARE\Classes\Interface\{000D0700-0000-0000-C000-000000000046}" -Recurse -Force
        if(!$testResults) {
            Write-Host -ForegroundColor Yellow "You're using Office Click2Run, so we need to fix some registry keys..."
            $registryKeyMods = '
            Copy-Item -Path "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Classes\CLSID\{00021A20-0000-0000-C000-000000000046}" -Destination "HKLM:\SOFTWARE\Classes\CLSID\{00021A20-0000-0000-C000-000000000046}" -Recurse -Force
            Copy-Item -Path "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Classes\Wow6432Node\CLSID\{00021A20-0000-0000-C000-000000000046}" -Destination "HKLM:\SOFTWARE\Classes\Wow6432Node\CLSID\{00021A20-0000-0000-C000-000000000046}" -Recurse -Force
            '
            $encodedCommand = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($registryKeyMods))
            Start-Process -FilePath powershell.exe -Verb runas -ArgumentList "-encodedCommand $encodedCommand" -wait
        }
    }
}

Export-ModuleMember Invoke-PatchOfficeC2RRegistry