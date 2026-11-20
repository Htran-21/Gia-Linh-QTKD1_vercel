param(
    [string]$TargetFile = "c:\Users\ADMIN\Downloads\AI Agent CK\scripts\build_enterprise_dataset.ps1"
)

$content = [System.IO.File]::ReadAllText($TargetFile, [System.Text.Encoding]::UTF8)
$utf8Bom = New-Object System.Text.UTF8Encoding($true)
[System.IO.File]::WriteAllText($TargetFile, $content, $utf8Bom)

$errors = $null
$tokens = [System.Management.Automation.PSParser]::Tokenize($content, [ref]$errors)
Write-Output ("Saved with UTF-8 BOM: " + $TargetFile)
Write-Output ("Errors count: " + $errors.Count)
