$ErrorActionPreference = "Stop"

$repos = @(
  @{
    Name = "cxy"
    Path = "C:\Users\admin\Documents\New project"
  },
  @{
    Name = "ui-cxy"
    Path = "E:\UI-cxy"
  },
  @{
    Name = "cp-cxy"
    Path = "E:\cxy"
  }
)

function Get-RepoStatus {
  param(
    [hashtable]$Repo
  )

  $repoPath = $Repo.Path
  $pidFile = Join-Path $repoPath ".git-auto-sync.pid"
  $logFile = Join-Path $repoPath ".git-auto-sync.log"
  $startScript = Join-Path $repoPath "start-auto-sync.ps1"

  $processState = "Stopped"
  $action = "None"
  $watcherPid = ""

  if (Test-Path -LiteralPath $pidFile) {
    $watcherPid = (Get-Content -LiteralPath $pidFile -ErrorAction SilentlyContinue | Select-Object -First 1).ToString().Trim()
    if ($watcherPid) {
      $process = Get-Process -Id $watcherPid -ErrorAction SilentlyContinue
      if ($process) {
        $processState = "Running"
      }
      else {
        Remove-Item -LiteralPath $pidFile -ErrorAction SilentlyContinue
      }
    }
  }

  if ($processState -ne "Running" -and (Test-Path -LiteralPath $startScript)) {
    & $startScript *> $null
    Start-Sleep -Seconds 2
    if (Test-Path -LiteralPath $pidFile) {
      $watcherPid = (Get-Content -LiteralPath $pidFile -ErrorAction SilentlyContinue | Select-Object -First 1).ToString().Trim()
      if ($watcherPid -and (Get-Process -Id $watcherPid -ErrorAction SilentlyContinue)) {
        $processState = "Running"
        $action = "Started"
      }
    }
  }

  $branchLine = git -C $repoPath status --short --branch 2>$null | Select-Object -First 1
  $changes = @(git -C $repoPath status --short 2>$null).Count

  $lastLog = ""
  if (Test-Path -LiteralPath $logFile) {
    $lastLog = Get-Content -LiteralPath $logFile -Tail 1 -ErrorAction SilentlyContinue
  }

  [pscustomobject]@{
    Repo = $Repo.Name
    Watcher = $processState
    Action = $action
    WatcherPid = $watcherPid
    Changes = $changes
    Branch = $branchLine
    LastLog = $lastLog
    Path = $repoPath
  }
}

$results = foreach ($repo in $repos) {
  Get-RepoStatus -Repo $repo
}

$results | Format-Table Repo,Watcher,Action,WatcherPid,Changes,Branch -AutoSize

Write-Host ""
foreach ($result in $results) {
  Write-Host ("[{0}] {1}" -f $result.Repo, $result.Path)
  if ($result.LastLog) {
    Write-Host ("  LastLog: {0}" -f $result.LastLog)
  }
}
