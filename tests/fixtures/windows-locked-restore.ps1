$ErrorActionPreference = 'Stop'

$worker = Join-Path $PSScriptRoot 'windows-restore-worker.cjs'
$root = Join-Path ([System.IO.Path]::GetTempPath()) ([System.Guid]::NewGuid().ToString('N'))
$document = Join-Path $root 'restore.docx'
New-Item -ItemType Directory -Path $root | Out-Null

function Invoke-Worker([string]$Mode) {
  $raw = & node $worker $Mode $document
  if ($LASTEXITCODE -ne 0 -and $Mode -ne 'restore') { throw "Worker $Mode failed: $raw" }
  return ($raw | ConvertFrom-Json)
}

try {
  $setup = Invoke-Worker 'setup'
  if (-not $setup.ok -or $setup.commits -ne 2) { throw "Setup did not create the expected history: $($setup | ConvertTo-Json -Compress)" }

  $locked = [System.IO.File]::Open($document, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::Read)
  try {
    $blocked = Invoke-Worker 'restore'
    if ($LASTEXITCODE -eq 0 -or $blocked.ok) { throw 'Restore unexpectedly succeeded while a real Windows handle denied delete sharing.' }
    $state = Invoke-Worker 'inspect'
    if (-not $state.ok -or $state.hash -ne $setup.editedHash -or $state.commits -ne 2) {
      throw "Locked restore changed document bytes or history: $($state | ConvertTo-Json -Compress)"
    }
    $backups = @(Get-ChildItem -LiteralPath $root -Force | Where-Object { $_.Name -like '*.astra-restore-backup-*' })
    if ($backups.Count -ne 0) { throw "Failed locked restore left unexpected backup(s): $($backups.Name -join ', ')" }
  }
  finally {
    $locked.Dispose()
  }

  $retried = Invoke-Worker 'restore'
  if ($LASTEXITCODE -ne 0 -or -not $retried.ok) { throw "Restore retry failed after releasing the Windows handle: $($retried | ConvertTo-Json -Compress)" }
  $restored = Invoke-Worker 'inspect'
  if (-not $restored.ok -or $restored.hash -ne $setup.originalHash -or $restored.commits -ne 3) {
    throw "Restore retry did not restore exact bytes and append one commit: $($restored | ConvertTo-Json -Compress)"
  }
  $remainingBackups = @(Get-ChildItem -LiteralPath $root -Force | Where-Object { $_.Name -like '*.astra-restore-backup-*' })
  if ($remainingBackups.Count -ne 0) { throw "Successful restore left unexpected backup(s): $($remainingBackups.Name -join ', ')" }
  Write-Output 'Windows locked-target restore gate passed.'
}
finally {
  Remove-Item -LiteralPath $root -Recurse -Force -ErrorAction SilentlyContinue
}
