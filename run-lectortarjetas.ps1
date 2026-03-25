Set-Location "C:\Users\teoja\OneDrive\Escritorio\Lectortarjetas"

while ($true) {
  try {
    "$(Get-Date -Format s) START uvicorn" | Out-File -FilePath ".\logs\launcher.log" -Append -Encoding utf8
    & ".\.venv\Scripts\python.exe" -m uvicorn app.main:app --host 127.0.0.1 --port 8001 >> ".\logs\app.log" 2>&1
    "$(Get-Date -Format s) EXIT CODE=$LASTEXITCODE" | Out-File -FilePath ".\logs\launcher.log" -Append -Encoding utf8
  } catch {
    "$(Get-Date -Format s) ERROR $($_.Exception.Message)" | Out-File -FilePath ".\logs\launcher.log" -Append -Encoding utf8
  }
  Start-Sleep -Seconds 5
}
