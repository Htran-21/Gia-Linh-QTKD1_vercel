$path = "c:\Users\ADMIN\Downloads\AI Agent CK\scripts\build_custom_lightblue_presentation.ps1"
$content = [System.IO.File]::ReadAllText($path, [System.Text.Encoding]::UTF8)

# Replace all occurrences of '", sz =' with '"; sz ='
$content = $content.Replace('", sz =', '"; sz =')

$utf8Bom = New-Object System.Text.UTF8Encoding($true)
[System.IO.File]::WriteAllText($path, $content, $utf8Bom)

# Check syntax
$errors = $null
$tokens = [System.Management.Automation.PSParser]::Tokenize($content, [ref]$errors)
Write-Output ("Parser Errors count: " + $errors.Count)
if ($errors.Count -gt 0) {
    foreach ($err in $errors) {
        Write-Output ("Line " + $err.Token.StartLine + ": " + $err.Message)
    }
} else {
    Write-Output "Syntax is 100% valid!"
}
