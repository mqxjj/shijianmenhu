try {
  $r = Invoke-WebRequest -Uri 'http://localhost:8765/market_regulation_portal.html' -TimeoutSec 5 -UseBasicParsing
  Write-Host "OK Status: $($r.StatusCode) Size: $($r.Content.Length)"
} catch {
  Write-Host "FAIL: $($_.Exception.Message)"
}
