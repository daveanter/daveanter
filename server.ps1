<#
  Shift Scheduler — self-contained PowerShell web server
  No Python, no installs required.
  Usage: right-click start.bat -> Run as administrator
  (Admin needed so other PCs on the network can reach the form)
#>
param([int]$Port = 5000)
Set-StrictMode -Off
$ErrorActionPreference = 'Continue'

# ── Paths ─────────────────────────────────────────────────────────────────────
$ROOT   = Split-Path -Parent $MyInvocation.MyCommand.Path
$CFG_F  = Join-Path $ROOT 'config.json'
$DATA_F = Join-Path $ROOT 'data\submissions.json'

if (!(Test-Path (Join-Path $ROOT 'data'))) {
    New-Item -ItemType Directory -Path (Join-Path $ROOT 'data') | Out-Null
}
if (!(Test-Path $DATA_F)) { '{"submissions":[]}' | Set-Content $DATA_F -Encoding UTF8 }

# ── Config ────────────────────────────────────────────────────────────────────
$CFG = Get-Content $CFG_F -Raw | ConvertFrom-Json

# ── Data helpers ──────────────────────────────────────────────────────────────
function Get-Data {
    try {
        $d = Get-Content $DATA_F -Raw | ConvertFrom-Json
        if ($null -eq $d.submissions) { $d.submissions = @() }
        return $d
    } catch { return [PSCustomObject]@{ submissions = @() } }
}

function Save-Data($d) {
    [PSCustomObject]@{ submissions = @($d.submissions) } |
        ConvertTo-Json -Depth 10 | Set-Content $DATA_F -Encoding UTF8
}

function New-Token {
    $b = New-Object byte[] 24
    ([Security.Cryptography.RandomNumberGenerator]::Create()).GetBytes($b)
    return [Convert]::ToBase64String($b) -replace '[+/=]', ''
}

# ── Time helpers ──────────────────────────────────────────────────────────────
function Fmt12h([int]$mins) {
    $h = [int]($mins / 60); $m = $mins % 60
    $suf = if ($h -lt 12) { 'AM' } else { 'PM' }
    $h12 = $h % 12; if ($h12 -eq 0) { $h12 = 12 }
    return "${h12}:$($m.ToString('D2')) $suf"
}

function Get-NextMonday {
    $today = [datetime]::Today
    $dow   = [int]$today.DayOfWeek
    $days  = if ($dow -eq 1) { 7 } else { (8 - $dow) % 7 }
    return $today.AddDays($days).ToString('yyyy-MM-dd')
}

function Get-LocalIP {
    if ($CFG.base_url) { return ($CFG.base_url -replace '^https?://', '' -replace ':.*', '') }
    try {
        $u = [Net.Sockets.UdpClient]::new(); $u.Connect('8.8.8.8', 80)
        $ip = $u.Client.LocalEndPoint.Address.ToString(); $u.Close(); return $ip
    } catch { return 'localhost' }
}

function Get-BaseUrl {
    if ($CFG.base_url) { return $CFG.base_url.TrimEnd('/') }
    return "http://$(Get-LocalIP):$Port"
}

# ── HTTP helpers ──────────────────────────────────────────────────────────────
function Parse-QS([string]$qs) {
    $d = @{}
    if ($qs -and $qs.StartsWith('?')) { $qs = $qs.Substring(1) }
    foreach ($p in ($qs -split '&')) {
        if ($p -match '^([^=]+)=(.*)$') {
            $d[[Uri]::UnescapeDataString($matches[1])] =
                [Uri]::UnescapeDataString($matches[2] -replace '\+', ' ')
        }
    }
    return $d
}

function Send-Html($ctx, [string]$html, [int]$code = 200) {
    $bytes = [Text.Encoding]::UTF8.GetBytes($html)
    $ctx.Response.StatusCode      = $code
    $ctx.Response.ContentType     = 'text/html; charset=utf-8'
    $ctx.Response.ContentLength64 = $bytes.Length
    $ctx.Response.OutputStream.Write($bytes, 0, $bytes.Length)
    $ctx.Response.OutputStream.Close()
}

function Send-Redirect($ctx, [string]$url) {
    $ctx.Response.Redirect($url); $ctx.Response.Close()
}

function Send-Bytes($ctx, [byte[]]$bytes, [string]$mime, [string]$fname) {
    $ctx.Response.StatusCode      = 200
    $ctx.Response.ContentType     = $mime
    $ctx.Response.ContentLength64 = $bytes.Length
    $ctx.Response.Headers.Add('Content-Disposition', "attachment; filename=`"$fname`"")
    $ctx.Response.OutputStream.Write($bytes, 0, $bytes.Length)
    $ctx.Response.OutputStream.Close()
}

# ── Sessions ──────────────────────────────────────────────────────────────────
$script:Sessions = @{}

function Get-Sid($ctx) {
    $c = $ctx.Request.Cookies['sess']; if ($c) { return $c.Value }; return $null
}

function Is-Mgr($ctx) {
    $s = Get-Sid $ctx; return ($s -and $script:Sessions.ContainsKey($s))
}

# ── Email ─────────────────────────────────────────────────────────────────────
function Send-Email([string]$to, [string]$toName, [string]$subj, [string]$html) {
    $smtp = New-Object Net.Mail.SmtpClient($CFG.smtp.host, [int]$CFG.smtp.port)
    $smtp.EnableSsl  = $true
    $smtp.Credentials = New-Object Net.NetworkCredential($CFG.smtp.username, $CFG.smtp.password)
    $msg = New-Object Net.Mail.MailMessage
    $msg.From = New-Object Net.Mail.MailAddress($CFG.smtp.from_email, $CFG.smtp.from_name)
    $msg.To.Add((New-Object Net.Mail.MailAddress($to, $toName)))
    $msg.Subject = $subj; $msg.IsBodyHtml = $true; $msg.Body = $html
    try { $smtp.Send($msg) } finally { $smtp.Dispose() }
}

function Email-Wrap([string]$sub, [string]$body) {
    return @"
<!DOCTYPE html><html><head><meta charset="utf-8"></head>
<body style="margin:0;padding:0;background:#f0f4f8;font-family:Calibri,'Segoe UI',Arial,sans-serif;">
<table width="100%" cellpadding="0" cellspacing="0" style="padding:28px 0;"><tr><td align="center">
<table width="600" cellpadding="0" cellspacing="0" style="background:#fff;border-radius:8px;box-shadow:0 2px 12px rgba(0,0,0,.09);">
<tr><td bgcolor="#1F4E79" style="padding:26px 36px;">
  <p style="margin:0;color:#fff;font-size:19px;font-weight:bold;">Weekly Shift Schedule</p>
  <p style="margin:5px 0 0;color:#9DC3E6;font-size:13px;">$sub</p>
</td></tr>
<tr><td style="padding:30px 36px;">$body</td></tr>
<tr><td bgcolor="#f7f9fb" style="padding:14px 36px;border-top:1px solid #e8ecf0;text-align:center;">
  <p style="margin:0;color:#bbb;font-size:11px;">Sent by your shift scheduler</p>
</td></tr></table></td></tr></table></body></html>
"@
}

function Notify-Mgr($sub) {
    $base   = Get-BaseUrl
    $wk     = [datetime]::ParseExact($sub.week_start, 'yyyy-MM-dd', $null)
    $wkE    = $wk.AddDays(4).ToString('MMMM dd, yyyy')
    $hrs    = [math]::Max(0, ($sub.end_mins - $sub.start_mins - $sub.lunch_mins) / 60)
    $appUrl = "$base/approve/$($sub.approve_token)"
    $body   = @"
<p style='font-size:16px;color:#333;margin:0 0 14px;'><strong>$($sub.name)</strong> submitted for <strong>$($wk.ToString('MMMM dd')) – $wkE</strong>.</p>
<table width='100%' cellpadding='0' cellspacing='0' style='border-collapse:collapse;margin:14px 0;'>
<tr><td style='padding:10px 14px;background:#EEF3F8;font-size:14px;border-bottom:1px solid #ddd;'><strong>Days:</strong> Monday – Friday</td></tr>
<tr><td style='padding:10px 14px;background:#fff;font-size:14px;border-bottom:1px solid #ddd;'><strong>Start:</strong> $(Fmt12h $sub.start_mins)</td></tr>
<tr><td style='padding:10px 14px;background:#EEF3F8;font-size:14px;border-bottom:1px solid #ddd;'><strong>Finish:</strong> $(Fmt12h $sub.end_mins)</td></tr>
<tr><td style='padding:10px 14px;background:#fff;font-size:14px;border-bottom:1px solid #ddd;'><strong>Lunch:</strong> $($sub.lunch_mins) min</td></tr>
<tr><td style='padding:10px 14px;background:#E2EFDA;font-size:14px;font-weight:bold;color:#375623;'>Daily: $($hrs.ToString('F1')) hrs &bull; Weekly: $([math]::Round($hrs*5,1)) hrs</td></tr>
</table>
<table cellpadding='0' cellspacing='0' style='margin:20px 0 0;'><tr>
<td bgcolor='#2E7D32' style='border-radius:6px;'>
<a href='$appUrl' style='display:inline-block;padding:13px 28px;color:#fff;font-size:15px;font-weight:bold;text-decoration:none;'>&#10003; Approve Schedule &rarr;</a>
</td></tr></table>
"@
    Send-Email $CFG.manager_email $CFG.manager_name `
        "Schedule: $($sub.name) — $($wk.ToString('MMM dd, yyyy'))" `
        (Email-Wrap "New schedule from $($sub.name)" $body)
}

# ── Time option lists ─────────────────────────────────────────────────────────
$START_OPTS = @()
for ($h = 6; $h -le 14; $h++) { foreach ($m in 0, 30) { $START_OPTS += [PSCustomObject]@{ mins = $h*60+$m; lbl = Fmt12h($h*60+$m) } } }
$END_OPTS = @()
for ($h = 11; $h -le 22; $h++) { foreach ($m in 0, 30) { $END_OPTS += [PSCustomObject]@{ mins = $h*60+$m; lbl = Fmt12h($h*60+$m) } } }

function Options-Html($opts, [int]$sel) {
    return ($opts | ForEach-Object {
        $s = if ($_.mins -eq $sel) { ' selected' } else { '' }
        "<option value='$($_.mins)'$s>$($_.lbl)</option>"
    }) -join ''
}

# ── Page builders ─────────────────────────────────────────────────────────────
function Page-Form([string]$week, [string]$err, [int]$s = 540, [int]$e = 1020, [int]$l = 30) {
    $wk  = [datetime]::ParseExact($week, 'yyyy-MM-dd', $null)
    $wkE = $wk.AddDays(4)
    $errH = if ($err) { "<div class='alert alert-warning py-2'>$err</div>" } else { '' }
    $l30  = if ($l -eq 30) { ' checked' } else { '' }
    $l60  = if ($l -eq 60) { ' checked' } else { '' }
    return @"
<!DOCTYPE html><html lang="en"><head>
<meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Submit Your Schedule</title>
<link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/css/bootstrap.min.css" rel="stylesheet">
<style>body{background:#f0f4f8;min-height:100vh;display:flex;align-items:center;}
.card{border:none;border-radius:12px;box-shadow:0 2px 20px rgba(0,0,0,.1);overflow:hidden;}
.card-header{background:#1F4E79;color:#fff;padding:22px 28px;}
.card-header h1{margin:0 0 3px;font-size:1.25rem;font-weight:700;}
.card-header p{margin:0;color:#9DC3E6;font-size:.85rem;}
.form-label{font-weight:700;color:#1F4E79;font-size:.82rem;margin-bottom:4px;}
.name-input{text-transform:uppercase;letter-spacing:.08em;font-weight:600;font-size:1.05rem;}
.btn-sub{background:#1F4E79;border:none;font-size:1rem;font-weight:600;padding:13px;}
.btn-sub:hover{background:#163859;}</style></head>
<body><div class="container py-4" style="max-width:460px;"><div class="card">
<div class="card-header"><h1>Submit Your Schedule</h1>
<p>$($wk.ToString('MMMM dd')) – $($wkE.ToString('MMMM dd, yyyy')) &bull; Monday – Friday</p></div>
<div class="card-body p-4">$errH
<form method="POST" action="/simple">
<input type="hidden" name="week" value="$week">
<div class="mb-3"><label class="form-label">YOUR FIRST NAME</label>
<input type="text" class="form-control name-input" name="name" placeholder="e.g. ALICE"
  maxlength="30" required oninput="this.value=this.value.toUpperCase()">
<div class="form-text">Enter your first name in capitals.</div></div>
<div class="mb-3"><label class="form-label">START TIME</label>
<select class="form-select" name="start_mins">$(Options-Html $START_OPTS $s)</select></div>
<div class="mb-3"><label class="form-label">FINISH TIME</label>
<select class="form-select" name="end_mins">$(Options-Html $END_OPTS $e)</select></div>
<div class="mb-4"><label class="form-label d-block">LUNCH BREAK</label>
<div class="d-flex gap-4">
<div class="form-check"><input class="form-check-input" type="radio" name="lunch_mins" value="30" id="l30"$l30><label class="form-check-label" for="l30">30 minutes</label></div>
<div class="form-check"><input class="form-check-input" type="radio" name="lunch_mins" value="60" id="l60"$l60><label class="form-check-label" for="l60">60 minutes</label></div>
</div></div>
<button type="submit" class="btn btn-primary btn-sub w-100">Send My Schedule &rarr;</button>
</form></div></div></div></body></html>
"@
}

function Page-Thanks([string]$name, [string]$week) {
    $wk = [datetime]::ParseExact($week, 'yyyy-MM-dd', $null)
    return @"
<!DOCTYPE html><html lang="en"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1"><title>Submitted</title>
<link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/css/bootstrap.min.css" rel="stylesheet">
<style>body{background:#f0f4f8;display:flex;align-items:center;min-height:100vh;}</style></head>
<body><div class="container text-center py-5" style="max-width:460px;">
<div style="font-size:3.5rem;">&#9989;</div>
<h2 class="mt-3 fw-bold" style="color:#1F4E79;">Thanks, $name!</h2>
<p class="text-muted mt-2">Your preferred schedule for the week of <strong>$($wk.ToString('MMMM dd, yyyy'))</strong> has been sent to your manager.</p>
<p class="text-muted" style="font-size:.85rem;">You can close this window.</p>
</div></body></html>
"@
}

function Page-Approved($sub) {
    $wk = [datetime]::ParseExact($sub.week_start, 'yyyy-MM-dd', $null)
    return @"
<!DOCTYPE html><html lang="en"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1"><title>Approved</title>
<link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/css/bootstrap.min.css" rel="stylesheet">
<style>body{background:#f0f4f8;display:flex;align-items:center;min-height:100vh;}</style></head>
<body><div class="container text-center py-5" style="max-width:460px;">
<div style="font-size:3.5rem;color:#2E7D32;">&#10003;</div>
<h2 class="mt-3 fw-bold" style="color:#2E7D32;">Approved!</h2>
<p class="text-muted mt-2"><strong>$($sub.name)'s</strong> schedule for the week of <strong>$($wk.ToString('MMMM dd, yyyy'))</strong> is confirmed.</p>
<div class="card border-0 shadow-sm mt-3 text-start" style="background:#E2EFDA;color:#375623;"><div class="card-body">
<div class="fw-bold mb-2">Monday – Friday</div>
<div>Start: <strong>$(Fmt12h $sub.start_mins)</strong></div>
<div>Finish: <strong>$(Fmt12h $sub.end_mins)</strong></div>
<div>Lunch: <strong>$($sub.lunch_mins) minutes</strong></div>
</div></div>
<p class="text-muted mt-3" style="font-size:.85rem;">You can close this window.</p>
</div></body></html>
"@
}

function Page-Error([string]$msg) {
    return @"
<!DOCTYPE html><html lang="en"><head><meta charset="utf-8"><title>Error</title>
<link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/css/bootstrap.min.css" rel="stylesheet">
<style>body{background:#f0f4f8;display:flex;align-items:center;min-height:100vh;}</style></head>
<body><div class="container text-center py-5" style="max-width:460px;">
<div class="alert alert-warning">$msg</div></div></body></html>
"@
}

function Page-Login([string]$err) {
    $errH = if ($err) { "<div class='alert alert-danger py-2'>$err</div>" } else { '' }
    return @"
<!DOCTYPE html><html lang="en"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1"><title>Manager Login</title>
<link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/css/bootstrap.min.css" rel="stylesheet">
<style>body{background:#f0f4f8;display:flex;align-items:center;min-height:100vh;}</style></head>
<body><div class="container py-4" style="max-width:380px;">
<div class="card border-0 shadow-sm" style="border-radius:12px;overflow:hidden;">
<div class="card-header" style="background:#1F4E79;color:#fff;padding:20px 24px;">
<h1 style="margin:0;font-size:1.2rem;font-weight:700;">Manager Login</h1></div>
<div class="card-body p-4">$errH
<form method="POST" action="/admin/login">
<div class="mb-3"><label class="form-label fw-bold" style="color:#1F4E79;font-size:.82rem;">PASSWORD</label>
<input type="password" class="form-control" name="password" required autofocus></div>
<button type="submit" class="btn btn-primary w-100 fw-bold" style="background:#1F4E79;border:none;">Log In</button>
</form></div></div></div></body></html>
"@
}

function Page-Dashboard([string]$selWeek) {
    $data  = Get-Data
    $weeks = @($data.submissions | ForEach-Object { $_.week_start } | Sort-Object -Unique -Descending)
    if (!$selWeek -and $weeks.Count -gt 0) { $selWeek = $weeks[0] }

    $wkBtns = if ($weeks.Count -gt 0) {
        ($weeks | ForEach-Object {
            $cls = if ($_ -eq $selWeek) { 'btn-primary' } else { 'btn-outline-secondary' }
            "<a href='/admin?week=$_' class='btn btn-sm $cls'>$_</a>"
        }) -join ' '
    } else { "<span class='text-muted'>No submissions yet. Download an invite .eml to get started.</span>" }

    $rows = ''
    if ($selWeek) {
        $subs = @($data.submissions | Where-Object { $_.week_start -eq $selWeek } | Sort-Object { $_.name })
        foreach ($s in $subs) {
            $hrs    = [math]::Max(0, ($s.end_mins - $s.start_mins - $s.lunch_mins) / 60)
            $badge  = if ($s.status -eq 'approved') { "<span class='badge bg-success'>Approved</span>" } else { "<span class='badge bg-secondary'>Pending</span>" }
            $rows  += "<tr><td class='ps-3 fw-bold' style='color:#1F4E79;'>$($s.name)</td>"
            $rows  += "<td>$(Fmt12h $s.start_mins)</td><td>$(Fmt12h $s.end_mins)</td>"
            $rows  += "<td>$($s.lunch_mins) min</td>"
            $rows  += "<td class='text-center fw-bold' style='background:#DAEEF3;'>$($hrs.ToString('F1')) h/day</td>"
            $rows  += "<td class='text-center'>$badge</td></tr>"
        }
        if (!$rows) { $rows = "<tr><td colspan='6' class='text-center text-muted py-3'>No submissions for this week.</td></tr>" }
    }

    $expBtn = if ($selWeek) { "<a href='/admin/export?week=$selWeek' class='btn btn-success'>&#11015; Export Excel</a>" } else { '' }
    $emlBtn = if ($selWeek) { "<a href='/admin/invite.eml?week=$selWeek' class='btn btn-outline-primary'>&#128233; Download Invite .eml</a>" } else { "<a href='/admin/invite.eml' class='btn btn-outline-primary'>&#128233; Download Invite .eml</a>" }
    $tbl    = if ($selWeek) { @"
<div class="bg-white rounded shadow-sm overflow-hidden mt-3">
<div class="table-responsive"><table class="table table-bordered table-sm mb-0">
<thead><tr><th class="ps-3">Employee</th><th>Start</th><th>Finish</th><th>Lunch</th><th class="text-center">Hours</th><th class="text-center">Status</th></tr></thead>
<tbody>$rows</tbody></table></div></div>
"@ } else { '' }

    return @"
<!DOCTYPE html><html lang="en"><head>
<meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Manager Dashboard</title>
<link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/css/bootstrap.min.css" rel="stylesheet">
<style>body{background:#f0f4f8;font-size:.88rem;}
.hdr{background:#1F4E79;color:#fff;padding:20px 28px;border-radius:10px;}
.hdr p{color:#9DC3E6;margin:4px 0 0;font-size:.85rem;}
th{background:#D6E4F0;color:#1F4E79;}</style></head>
<body><div class="container-fluid py-4" style="max-width:1000px;">
<div class="d-flex justify-content-between align-items-start mb-3 flex-wrap gap-2">
<div class="hdr flex-grow-1"><h1 style="font-size:1.3rem;font-weight:700;margin:0;">Manager Dashboard</h1>
<p>Employee Weekly Shift Schedules</p></div>
<div class="d-flex gap-2 ms-3 align-items-start flex-wrap">
$expBtn $emlBtn
<a href='/admin/logout' class='btn btn-outline-secondary'>Log out</a></div></div>
<div class="d-flex gap-2 flex-wrap">$wkBtns</div>
$tbl
</div></body></html>
"@
}

# ── Excel timeline export ─────────────────────────────────────────────────────
function Export-Timeline([string]$week) {
    $data = Get-Data
    $subs = @($data.submissions | Where-Object { $_.week_start -eq $week } | Sort-Object { $_.name })
    $wkDt = [datetime]::ParseExact($week, 'yyyy-MM-dd', $null)

    $xl = $null; $wb = $null
    try {
        $xl = New-Object -ComObject Excel.Application
        $xl.Visible = $false; $xl.DisplayAlerts = $false
        $wb = $xl.Workbooks.Add()
        $ws = $wb.Worksheets.Item(1)
        $ws.Name = 'Weekly Timeline'

        # slot setup: 8:00–21:00 in 30-min steps = 26 slots
        $slots = @(); for ($m = 480; $m -lt 1260; $m += 30) { $slots += $m }
        $nS = $slots.Count  # 26
        $CL = 1; $CS = 2; $CR = $CS + $nS; $CH = $CR + 1

        $ws.Columns.Item($CL).ColumnWidth = 20
        for ($c = $CS; $c -lt $CS + $nS; $c++) { $ws.Columns.Item($c).ColumnWidth = 2.6 }
        $ws.Columns.Item($CR).ColumnWidth = 20
        $ws.Columns.Item($CH).ColumnWidth = 7

        function XC([string]$h) {
            $r=[Convert]::ToInt32($h.Substring(0,2),16)
            $g=[Convert]::ToInt32($h.Substring(2,2),16)
            $b=[Convert]::ToInt32($h.Substring(4,2),16)
            return $b*65536+$g*256+$r
        }
        function SC($cell,[string]$bg,[string]$fg='000000',[bool]$bold=$false,[int]$sz=10,[int]$ha=-4108) {
            $cell.Interior.Color=$cell.Font.Color=0
            $cell.Interior.Color=XC $bg; $cell.Font.Color=XC $fg
            $cell.Font.Bold=$bold; $cell.Font.Size=$sz; $cell.Font.Name='Calibri'
            $cell.HorizontalAlignment=$ha; $cell.VerticalAlignment=-4108
        }

        $row = 1
        # Title
        $ws.Rows.Item($row).RowHeight = 28
        $ws.Range($ws.Cells($row,1),$ws.Cells($row,$CH)).Merge() | Out-Null
        $c = $ws.Cells($row,1)
        $c.Value2 = "Staff Coverage Timeline  —  Week of $($wkDt.ToString('MMMM dd, yyyy'))  |  Mon–Fri  |  8 AM–9 PM  |  30-min slots"
        SC $c '1F4E79' 'FFFFFF' $true 13 -4131
        $row++

        # Header row
        $ws.Rows.Item($row).RowHeight = 24
        foreach ($h in @(@{c=$CL;v='Employee';b='D6E4F0';f='1F4E79'},@{c=$CR;v='Employee';b='D6E4F0';f='1F4E79'},@{c=$CH;v='Day Hrs';b='E2EFDA';f='375623'})) {
            $c = $ws.Cells($row,$h.c); $c.Value2 = $h.v; SC $c $h.b $h.f $true 10
        }
        for ($si = 0; $si -lt $nS; $si++) {
            $mn = $slots[$si]; $hh = [int]($mn/60); $mm = $mn%60
            $c = $ws.Cells($row,$CS+$si)
            $c.Value2 = if ($mm -eq 0) { "$hh" } else { ':30' }
            $c.Interior.Color = XC 'EBF3FB'; $c.Font.Color = XC '1F4E79'
            $c.Font.Size = 7; $c.Font.Bold = ($mm -eq 0); $c.Font.Name = 'Calibri'
            $c.HorizontalAlignment = -4108; $c.VerticalAlignment = 4; $c.Orientation = 90
        }
        $row++

        # Day sections
        $DAYS = 'Monday','Tuesday','Wednesday','Thursday','Friday'
        foreach ($di in 0..4) {
            $dayDt = $wkDt.AddDays($di)
            $ws.Rows.Item($row).RowHeight = 20
            $ws.Range($ws.Cells($row,1),$ws.Cells($row,$CH)).Merge() | Out-Null
            $c = $ws.Cells($row,1)
            $c.Value2 = "  $($dayDt.ToString('dddd, MMMM dd, yyyy').ToUpper())"
            SC $c '1F4E79' 'FFFFFF' $true 11 -4131
            $row++

            if ($subs.Count -gt 0) {
                for ($ei = 0; $ei -lt $subs.Count; $ei++) {
                    $s  = $subs[$ei]
                    $bg = if ($ei % 2 -eq 0) { 'EEF3F8' } else { 'F7FAFB' }
                    $ws.Rows.Item($row).RowHeight = 15

                    $sm = $s.start_mins; $em = $s.end_mins; $lm = $s.lunch_mins
                    $hrs = [math]::Max(0, ($em - $sm - $lm) / 60)
                    $mid = [int](($sm + $em) / 2)
                    $ls  = $mid - [int]($lm / 2)
                    $le  = $ls + $lm

                    foreach ($col in $CL, $CR) {
                        $c = $ws.Cells($row,$col); $c.Value2 = $s.name; SC $c $bg '1F4E79' $true 9 -4131
                    }
                    for ($si = 0; $si -lt $nS; $si++) {
                        $slot   = $slots[$si]
                        $slotBg = $bg
                        if ($slot -ge $sm -and $slot -lt $em) {
                            $slotBg = if ($slot -ge $ls -and $slot -lt $le) { 'FFEB9C' } else { '4472C4' }
                        }
                        $ws.Cells($row,$CS+$si).Interior.Color = XC $slotBg
                    }
                    $c = $ws.Cells($row,$CH)
                    if ($hrs -gt 0) { $c.Value2=[math]::Round($hrs,1); SC $c 'E2EFDA' '375623' $true 9 }
                    else            { $c.Value2='—'; SC $c 'F2F2F2' '888888' $false 9 }
                    $row++
                }
            }
            $ws.Rows.Item($row).RowHeight = 5; $row++
        }

        # Weekly summary
        $row++
        $ws.Rows.Item($row).RowHeight = 22
        $ws.Range($ws.Cells($row,1),$ws.Cells($row,$CH)).Merge() | Out-Null
        $c = $ws.Cells($row,1); $c.Value2 = '  Weekly Hours Summary'; SC $c '1F4E79' 'FFFFFF' $true 11 -4131
        $row++
        $ws.Rows.Item($row).RowHeight = 17
        $ws.Range($ws.Cells($row,1),$ws.Cells($row,$CR)).Merge() | Out-Null
        $c = $ws.Cells($row,1); $c.Value2 = 'Employee'; SC $c 'D6E4F0' '1F4E79' $true 10
        $c = $ws.Cells($row,$CH); $c.Value2 = 'Total Hrs'; SC $c 'E2EFDA' '375623' $true 10
        $row++

        if ($subs.Count -gt 0) {
            for ($ei = 0; $ei -lt $subs.Count; $ei++) {
                $s  = $subs[$ei]; $bg = if ($ei%2-eq 0){'EEF3F8'}else{'F7FAFB'}
                $tot = [math]::Max(0,($s.end_mins-$s.start_mins-$s.lunch_mins)/60)*5
                $ws.Rows.Item($row).RowHeight = 16
                $ws.Range($ws.Cells($row,1),$ws.Cells($row,$CR)).Merge() | Out-Null
                $c = $ws.Cells($row,1); $c.Value2=$s.name; SC $c $bg '000000' $true 11 -4131
                $c = $ws.Cells($row,$CH); $c.Value2=[math]::Round($tot,1); SC $c 'C6EFCE' '375623' $true 11
                $row++
            }
        }

        # Legend
        $row++
        $ws.Cells($row,1).Value2='Legend:'; $ws.Cells($row,1).Font.Bold=$true; $ws.Cells($row,1).Font.Size=9
        $leg=@(@('4472C4','On Shift'),@('FFEB9C','Lunch Break'),@('EEF3F8','Off / Not Working'))
        for ($i=0;$i-lt$leg.Count;$i++) {
            $ws.Cells($row,3+$i*2).Interior.Color=XC $leg[$i][0]
            $ws.Cells($row,4+$i*2).Value2="  $($leg[$i][1])"
        }

        $xl.ActiveWindow.SplitRow    = 2
        $xl.ActiveWindow.SplitColumn = 1
        $xl.ActiveWindow.FreezePanes = $true
        $ws.PageSetup.Orientation    = 2   # landscape
        $ws.PageSetup.FitToPagesWide = 1
        $ws.PageSetup.FitToPagesTall = $false

        $tmp = [IO.Path]::Combine([IO.Path]::GetTempPath(), "timeline_$week.xlsx")
        $wb.SaveAs($tmp, 51)   # 51 = xlOpenXMLWorkbook
        return $tmp
    } finally {
        if ($wb)  { $wb.Close($false) }
        if ($xl)  { $xl.Quit(); [Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null }
        [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    }
}

# ── .eml builder ─────────────────────────────────────────────────────────────
function Build-Eml([string]$week) {
    $wk   = [datetime]::ParseExact($week,'yyyy-MM-dd',$null)
    $wkE  = $wk.AddDays(4)
    $url  = "$(Get-BaseUrl)/simple?week=$week"
    $btn  = "<table cellpadding='0' cellspacing='0' style='margin:20px 0 0;'><tr><td bgcolor='#1F4E79' style='border-radius:6px;'><a href='$url' style='display:inline-block;padding:13px 28px;color:#fff;font-size:15px;font-weight:bold;text-decoration:none;'>Submit My Schedule &rarr;</a></td></tr></table>"
    $body = @"
<p style='margin:0 0 14px;font-size:16px;color:#333;'>Hi,</p>
<p style='margin:0 0 18px;font-size:15px;line-height:1.6;color:#555;'>Please submit your preferred hours for the week of <strong>$($wk.ToString('MMMM dd')) – $($wkE.ToString('MMMM dd, yyyy'))</strong>.</p>
<p style='margin:0 0 22px;font-size:14px;line-height:1.6;color:#777;'>It only takes a minute — your name, start time, finish time, and lunch break.</p>
$btn
<p style='margin:22px 0 0;font-size:12px;color:#bbb;'>Please submit by end of day Thursday.</p>
"@
    $html = Email-Wrap "Week of $($wk.ToString('MMMM dd, yyyy'))" $body
    $date = [datetime]::Now.ToString('ddd, dd MMM yyyy HH:mm:ss zzz')
    $eml  = "MIME-Version: 1.0`r`nDate: $date`r`nFrom: $($CFG.manager_name) <$($CFG.manager_email)>`r`nTo: `r`nSubject: Your Schedule — Week of $($wk.ToString('MMMM dd, yyyy'))`r`nContent-Type: text/html; charset=`"UTF-8`"`r`n`r`n$html"
    return [Text.Encoding]::UTF8.GetBytes($eml)
}

# ── Router ────────────────────────────────────────────────────────────────────
function Handle($ctx) {
    $req  = $ctx.Request
    $path = $req.Url.LocalPath.TrimEnd('/')
    $meth = $req.HttpMethod
    $qs   = Parse-QS $req.Url.Query
    $form = @{}
    if ($meth -eq 'POST') {
        $rd = [IO.StreamReader]::new($req.InputStream, $req.ContentEncoding)
        $form = Parse-QS ("?" + $rd.ReadToEnd()); $rd.Close()
    }

    try {
        if ($path -eq '' -or $path -eq '/') { Send-Redirect $ctx '/simple'; return }

        # ── Employee form ──────────────────────────────────────────────────────
        if ($path -eq '/simple') {
            $week = if ($meth -eq 'POST') { $form['week'] } else { $qs['week'] }
            if (!$week) { $week = Get-NextMonday }

            if ($meth -eq 'POST') {
                $name = ($form['name']).Trim().ToUpper()
                $sMin = [int]$(if ($form['start_mins']) { $form['start_mins'] } else { 540 })
                $eMin = [int]$(if ($form['end_mins'])   { $form['end_mins']   } else { 1020 })
                $lMin = [int]$(if ($form['lunch_mins']) { $form['lunch_mins'] } else { 30 })
                if (!$name) { Send-Html $ctx (Page-Form $week 'Please enter your first name.' $sMin $eMin $lMin); return }

                $tok = New-Token
                $d   = Get-Data
                $sub = [PSCustomObject]@{
                    name          = $name
                    week_start    = $week
                    start_mins    = $sMin
                    end_mins      = $eMin
                    lunch_mins    = $lMin
                    approve_token = $tok
                    status        = 'pending'
                    submitted_at  = [datetime]::Now.ToString('o')
                }
                $d.submissions = @($d.submissions) + $sub
                Save-Data $d
                try { Notify-Mgr $sub } catch { Write-Warning "Email: $_" }
                Send-Html $ctx (Page-Thanks $name $week)
            } else {
                $sMin = [int]$(if ($qs['start_mins']) { $qs['start_mins'] } else { 540 })
                $eMin = [int]$(if ($qs['end_mins'])   { $qs['end_mins']   } else { 1020 })
                $lMin = [int]$(if ($qs['lunch_mins']) { $qs['lunch_mins'] } else { 30 })
                Send-Html $ctx (Page-Form $week '' $sMin $eMin $lMin)
            }
            return
        }

        # ── Approval ───────────────────────────────────────────────────────────
        if ($path -match '^/approve/(.+)$') {
            $tok = $matches[1]
            $d   = Get-Data
            $sub = $d.submissions | Where-Object { $_.approve_token -eq $tok } | Select-Object -First 1
            if (!$sub) { Send-Html $ctx (Page-Error 'This approval link is no longer valid.'); return }
            $sub.status = 'approved'
            Save-Data $d
            Send-Html $ctx (Page-Approved $sub)
            return
        }

        # ── Admin login/logout ─────────────────────────────────────────────────
        if ($path -eq '/admin/login') {
            if ($meth -eq 'POST') {
                if ($form['password'] -eq $CFG.manager_password) {
                    $sid = New-Token; $script:Sessions[$sid] = $true
                    $ctx.Response.SetCookie((New-Object Net.Cookie('sess',$sid,'/')))
                    Send-Redirect $ctx '/admin'
                } else { Send-Html $ctx (Page-Login 'Incorrect password.') }
            } else { Send-Html $ctx (Page-Login '') }
            return
        }

        if ($path -eq '/admin/logout') {
            $sid = Get-Sid $ctx; if ($sid) { $script:Sessions.Remove($sid) }
            Send-Redirect $ctx '/admin/login'; return
        }

        if (!(Is-Mgr $ctx)) { Send-Redirect $ctx '/admin/login'; return }

        # ── Admin dashboard ────────────────────────────────────────────────────
        if ($path -eq '/admin') {
            Send-Html $ctx (Page-Dashboard $(if ($qs['week']) { $qs['week'] } else { '' })); return
        }

        # ── Excel export ───────────────────────────────────────────────────────
        if ($path -eq '/admin/export') {
            $week = if ($qs['week']) { $qs['week'] } else { Get-NextMonday }
            try {
                $tmp   = Export-Timeline $week
                $bytes = [IO.File]::ReadAllBytes($tmp)
                Remove-Item $tmp -ErrorAction SilentlyContinue
                Send-Bytes $ctx $bytes 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' "timeline_$week.xlsx"
            } catch { Send-Html $ctx (Page-Error "Excel export failed: $_") }
            return
        }

        # ── Invite .eml ────────────────────────────────────────────────────────
        if ($path -eq '/admin/invite.eml') {
            $week = if ($qs['week']) { $qs['week'] } else { Get-NextMonday }
            Send-Bytes $ctx (Build-Eml $week) 'message/rfc822' "invite_$week.eml"
            return
        }

        Send-Html $ctx (Page-Error 'Page not found.') 404

    } catch {
        Write-Warning "Error handling $($req.Url): $_"
        try { Send-Html $ctx (Page-Error "Server error: $_") 500 } catch {}
    }
}

# ── Start ─────────────────────────────────────────────────────────────────────
$listener = New-Object System.Net.HttpListener
$networkOk = $false
try {
    $listener.Prefixes.Add("http://+:$Port/")
    $listener.Start()
    $networkOk = $true
} catch {
    Write-Warning "Network listener failed — falling back to localhost only."
    Write-Warning "Right-click start.bat and choose 'Run as administrator' for network access."
    $listener = New-Object System.Net.HttpListener
    $listener.Prefixes.Add("http://localhost:$Port/")
    $listener.Start()
}

$ip = Get-LocalIP
Write-Host ""
if ($networkOk) {
    Write-Host "  Shift Scheduler is running!" -ForegroundColor Green
    Write-Host "  Employee form : http://${ip}:$Port/simple"  -ForegroundColor Cyan
    Write-Host "  Manager login : http://${ip}:$Port/admin"   -ForegroundColor Cyan
} else {
    Write-Host "  Running on this PC only (localhost)" -ForegroundColor Yellow
    Write-Host "  Employee form : http://localhost:$Port/simple" -ForegroundColor Cyan
    Write-Host "  Manager login : http://localhost:$Port/admin"  -ForegroundColor Cyan
}
Write-Host "  Press Ctrl+C to stop." -ForegroundColor Gray
Write-Host ""

try {
    while ($listener.IsListening) {
        $ctx = $listener.GetContext()
        Handle $ctx
    }
} finally {
    $listener.Stop()
}
