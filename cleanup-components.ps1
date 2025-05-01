# PowerShell script to safely clean up redundant components
# Create backup directory if it doesn't exist
$backupDir = "C:\Users\shrey\OfficeAddinApps\core-models-excel-addin\backup-components"
if (-not (Test-Path $backupDir)) {
    New-Item -ItemType Directory -Path $backupDir | Out-Null
    Write-Host "Created backup directory: $backupDir"
}

# Components to keep in taskpane/components
$keepComponents = @("App.tsx", "Header.tsx")

# Back up and remove redundant components from taskpane/components
$taskpaneComponentsDir = "C:\Users\shrey\OfficeAddinApps\core-models-excel-addin\src\taskpane\components"
$componentsToRemove = Get-ChildItem -Path $taskpaneComponentsDir -Filter "*.tsx" | Where-Object { $keepComponents -notcontains $_.Name }

foreach ($component in $componentsToRemove) {
    # Create backup
    $backupPath = Join-Path -Path $backupDir -ChildPath $component.Name
    Copy-Item -Path $component.FullName -Destination $backupPath
    Write-Host "Backed up: $($component.Name) to $backupPath"
    
    # Remove the component
    Remove-Item -Path $component.FullName
    Write-Host "Removed redundant component: $($component.Name)"
}

Write-Host "`nCleanup completed successfully!"
Write-Host "Kept components in taskpane/components: $($keepComponents -join ', ')"
Write-Host "All removed components have been backed up to: $backupDir"
Write-Host "If you need to restore any components, you can find them in the backup directory."
