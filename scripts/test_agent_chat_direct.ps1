. "c:\Users\ADMIN\Downloads\AI Agent CK\server.ps1"
try {
    $res = Invoke-AgentChat -UserMessage "khác biệt null và 0" -LlmProvider "gemini" -LlmModel "gemini-2.5-flash" -ApiKey ""
    Write-Output "SUCCESS:"
    Write-Output $res
} catch {
    Write-Error ("Exception: " + $_.Exception.ToString())
}
