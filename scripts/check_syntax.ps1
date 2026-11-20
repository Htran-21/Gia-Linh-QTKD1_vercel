$path = "c:\Users\ADMIN\Downloads\AI Agent CK\server.ps1"
$content = [System.IO.File]::ReadAllText($path, [System.Text.Encoding]::UTF8)
$errors = $null
$tokens = [System.Management.Automation.PSParser]::Tokenize($content, [ref]$errors)
Write-Output ("Errors count: " + $errors.Count)
if ($errors.Count -gt 0) {
    foreach ($err in $errors) {
        Write-Output ("Line " + $err.Token.StartLine + ": " + $err.Message)
    }
} else {
    Write-Output "Syntax valid! No errors found."
}
