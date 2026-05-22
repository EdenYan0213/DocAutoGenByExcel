param(
    [string]$OutputDir = "",
    [string]$Title = "DocAutoGenByExcel &#24615;&#33021;&#32467;&#26524;&#35777;&#26126;",
    [int]$WindowWidth = 1600,
    [int]$WindowHeight = 980
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
$scriptTimer = [System.Diagnostics.Stopwatch]::StartNew()

function Resolve-WorkspaceRoot {
    return (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
}

function Resolve-EdgePath {
    $candidates = @()

    if ($env:ProgramFiles) {
        $candidates += (Join-Path $env:ProgramFiles "Microsoft\Edge\Application\msedge.exe")
    }
    if (${env:ProgramFiles(x86)}) {
        $candidates += (Join-Path ${env:ProgramFiles(x86)} "Microsoft\Edge\Application\msedge.exe")
    }

    foreach ($p in $candidates) {
        if ($p -and (Test-Path $p)) {
            return $p
        }
    }

    return $null
}

function New-BarRowHtml {
    param(
        [string]$Name,
        [double]$Seconds,
        [double]$Percent
    )

    return @"
    <div class=\"bar-row\">
      <div class=\"bar-label\">$Name</div>
      <div class=\"bar-track\"><div class=\"bar\" style=\"width: $Percent%;\"></div></div>
      <div class=\"bar-value\">$Seconds s</div>
    </div>
"@
}

$workspaceRoot = Resolve-WorkspaceRoot
if ([string]::IsNullOrWhiteSpace($OutputDir)) {
    $OutputDir = Join-Path $workspaceRoot "output\benchmark-proof"
}

if (-not (Test-Path $OutputDir)) {
    New-Item -ItemType Directory -Path $OutputDir | Out-Null
}

$createdAt = Get-Date
$createdAtText = $createdAt.ToString("yyyy-MM-dd HH:mm:ss")
$dateForFile = $createdAt.ToString("yyyyMMdd_HHmmss")

# User provided benchmark table
$rows = @(
    [PSCustomObject]@{
        Scale = "&#23567;&#35268;&#27169;"
        DataDesc = "2&#26465; / 2&#27169;&#22359; / 8&#21015;"
        Seconds = 3.13
        Note = "&#22522;&#32447;&#32791;&#26102;&#65292;&#20027;&#35201;&#21463;JVM&#20919;&#21551;&#21160;&#24433;&#21709;"
    },
    [PSCustomObject]@{
        Scale = "&#20013;&#35268;&#27169;"
        DataDesc = "6&#26465; / 4&#27169;&#22359; / 14&#21015;"
        Seconds = 3.36
        Note = "&#31456;&#33410;&#19982;&#34920;&#26684;&#22635;&#20805;&#22686;&#21152;&#65292;&#32791;&#26102;&#23567;&#24133;&#19978;&#21319;"
    },
    [PSCustomObject]@{
        Scale = "&#22823;&#35268;&#27169;"
        DataDesc = "35&#26465; / 2&#27169;&#22359; / 17&#21015;"
        Seconds = 3.91
        Note = "&#23376;&#31456;&#33410;&#19982;&#30446;&#24405;&#26356;&#26032;&#26174;&#33879;&#22686;&#22810;&#65292;&#25972;&#20307;&#20173;&#22312;&#21487;&#25509;&#21463;&#33539;&#22260;"
    }
)

$maxSeconds = ($rows | Measure-Object -Property Seconds -Maximum).Maximum
$minSeconds = ($rows | Measure-Object -Property Seconds -Minimum).Minimum

$tableRowsHtml = ($rows | ForEach-Object {
    "<tr><td>$($_.Scale)</td><td>$($_.DataDesc)</td><td class='num'>$('{0:N2}' -f $_.Seconds)</td><td>$($_.Note)</td></tr>"
}) -join [Environment]::NewLine

$barRowsHtml = ($rows | ForEach-Object {
    $pct = [Math]::Round(($_.Seconds / $maxSeconds) * 100, 1)
    New-BarRowHtml -Name $_.Scale -Seconds $_.Seconds -Percent $pct
}) -join [Environment]::NewLine

$summaryText = "Min $('{0:N2}' -f $minSeconds)s, Max $('{0:N2}' -f $maxSeconds)s, Delta $('{0:N2}' -f ($maxSeconds - $minSeconds))s"

$htmlPath = Join-Path $OutputDir "benchmark_proof_$dateForFile.html"
$pngPath = Join-Path $OutputDir "benchmark_proof_$dateForFile.png"
$csvPath = Join-Path $OutputDir "benchmark_data_$dateForFile.csv"

$html = @"
<!doctype html>
<html lang="en">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>$Title</title>
  <style>
    :root {
      --bg1: #f7f4ec;
      --bg2: #f2ece0;
      --card: #fffaf0;
      --text: #1f2937;
      --muted: #6b7280;
      --line: #d1d5db;
      --accent: #0f766e;
      --accent2: #115e59;
    }

    * { box-sizing: border-box; }
    body {
      margin: 0;
      font-family: "Segoe UI", Arial, sans-serif;
      color: var(--text);
      background:
        radial-gradient(circle at 5% 10%, rgba(15,118,110,0.15), transparent 35%),
        radial-gradient(circle at 95% 85%, rgba(180,83,9,0.12), transparent 40%),
        linear-gradient(145deg, var(--bg1), var(--bg2));
      min-height: 100vh;
      padding: 28px;
    }

    .wrap {
      max-width: 1320px;
      margin: 0 auto;
      background: var(--card);
      border: 1px solid var(--line);
      border-radius: 18px;
      box-shadow: 0 10px 30px rgba(0,0,0,0.08);
      overflow: hidden;
    }

    .head {
      padding: 20px 24px 16px;
      border-bottom: 1px solid var(--line);
      background: linear-gradient(90deg, rgba(15,118,110,0.12), rgba(245,158,11,0.12));
    }

    .title {
      margin: 0;
      font-size: 30px;
      letter-spacing: 0.5px;
    }

    .meta {
      margin-top: 8px;
      color: var(--muted);
      font-size: 14px;
    }

    .content {
      display: grid;
      grid-template-columns: 1.2fr 0.8fr;
      gap: 18px;
      padding: 20px;
    }

    .panel {
      border: 1px solid var(--line);
      border-radius: 14px;
      padding: 14px;
      background: #fff;
    }

    h2 {
      margin: 0 0 10px;
      font-size: 20px;
    }

    table {
      width: 100%;
      border-collapse: collapse;
      font-size: 15px;
    }

    th, td {
      border: 1px solid var(--line);
      padding: 10px;
      vertical-align: top;
    }

    th {
      background: #f0fdfa;
      text-align: left;
      font-weight: 700;
    }

    td.num {
      text-align: right;
      font-weight: 700;
      color: var(--accent2);
    }

    .bar-row {
      display: grid;
      grid-template-columns: 88px 1fr 64px;
      align-items: center;
      gap: 8px;
      margin: 10px 0;
      font-size: 14px;
    }

    .bar-label { font-weight: 700; }
    .bar-track {
      height: 18px;
      border: 1px solid #cbd5e1;
      border-radius: 999px;
      overflow: hidden;
      background: #f8fafc;
    }

    .bar {
      height: 100%;
      background: linear-gradient(90deg, var(--accent), #14b8a6);
    }

    .bar-value {
      text-align: right;
      font-weight: 700;
      color: #0f172a;
    }

    .summary {
      margin-top: 12px;
      color: #374151;
      font-size: 14px;
    }

    .foot {
      padding: 10px 20px 18px;
      color: var(--muted);
      font-size: 12px;
    }

    @media (max-width: 1000px) {
      .content { grid-template-columns: 1fr; }
      .title { font-size: 24px; }
    }
  </style>
</head>
<body>
  <div class="wrap">
    <div class="head">
      <h1 class="title">$Title</h1>
      <div class="meta">Generated at: $createdAtText | Workspace: $workspaceRoot</div>
    </div>

    <div class="content">
      <section class="panel">
        <h2>&#24615;&#33021;&#27979;&#37327;&#34920;</h2>
        <table>
          <thead>
            <tr>
              <th>&#35268;&#27169;</th>
              <th>&#25968;&#25454;&#37327;</th>
              <th>&#32791;&#26102;(s)</th>
              <th>&#35828;&#26126;</th>
            </tr>
          </thead>
          <tbody>
            $tableRowsHtml
          </tbody>
        </table>
      </section>

      <section class="panel">
        <h2>&#32791;&#26102;&#23545;&#27604;</h2>
        $barRowsHtml
        <div class="summary">$summaryText</div>
      </section>
    </div>

    <div class="foot">
      &#30001;&#33050;&#26412;&#33258;&#21160;&#29983;&#25104;&#65292;&#21487;&#30452;&#25509;&#29992;&#20110;&#35770;&#25991;&#25110;&#27719;&#25253;&#30340;&#25130;&#22270;&#35777;&#26126;&#12290;
    </div>
  </div>
</body>
</html>
"@

Set-Content -Path $htmlPath -Value $html -Encoding UTF8
$rows | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8

$edgePath = Resolve-EdgePath
$screenshotSuccess = $false
if ($edgePath) {
    $fileUri = "file:///" + ($htmlPath -replace "\\", "/")
    $edgeArgs = @(
        "--headless=new",
        "--disable-gpu",
        "--hide-scrollbars",
        "--window-size=$WindowWidth,$WindowHeight",
        "--screenshot=$pngPath",
        $fileUri
    )

    $proc = Start-Process -FilePath $edgePath -ArgumentList $edgeArgs -Wait -PassThru -WindowStyle Hidden
    if ($proc.ExitCode -eq 0 -and (Test-Path $pngPath)) {
        $screenshotSuccess = $true
    }
}

$scriptTimer.Stop()
$avgSeconds = [Math]::Round((($rows | Measure-Object -Property Seconds -Average).Average), 2)
$deltaSeconds = [Math]::Round(($maxSeconds - $minSeconds), 2)

Write-Host ""
Write-Host "========== Benchmark Timing Statistics =========="
Write-Host ("{0,-10} {1,-24} {2,8}  {3}" -f "Scale", "Data Size", "Time(s)", "Note")
foreach ($row in $rows) {
  Write-Host ("{0,-10} {1,-24} {2,8:N2}  {3}" -f $row.Scale, $row.DataDesc, $row.Seconds, $row.Note)
}
Write-Host "-------------------------------------------------"
Write-Host ("Min: {0:N2}s" -f $minSeconds)
Write-Host ("Max: {0:N2}s" -f $maxSeconds)
Write-Host ("Avg: {0:N2}s" -f $avgSeconds)
Write-Host ("Delta: {0:N2}s" -f $deltaSeconds)
Write-Host ("Script elapsed: {0}" -f $scriptTimer.Elapsed.ToString())
Write-Host "================================================="

Write-Host "[OK] HTML evidence page: $htmlPath"
Write-Host "[OK] CSV data:          $csvPath"
if ($screenshotSuccess) {
    Write-Host "[OK] PNG screenshot:    $pngPath"
} else {
    Write-Warning "PNG auto-shot not generated (Edge might be unavailable). Open the HTML page and capture manually."
}
