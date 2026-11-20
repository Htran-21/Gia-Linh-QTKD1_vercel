$tests = @("DN-2026-004", "DN-2026-007", "khác biệt null và 0")
foreach ($q in $tests) {
    Write-Output "========================================"
    Write-Output "TEST: $q"
    $body = @{ message = $q } | ConvertTo-Json
    $res = Invoke-RestMethod -Uri "http://localhost:8080/api/chat" -Method Post -ContentType "application/json; charset=utf-8" -Body ([System.Text.Encoding]::UTF8.GetBytes($body))
    Write-Output $res.reply
}
