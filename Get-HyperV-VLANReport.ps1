param(
    [switch]$Compact,
    [string]$DnsSuffix
)

$rows = @()

# ---------------------------
# 0) Pre-fetch objects and set up progress bar
# ---------------------------
$vms = Get-VM -ErrorAction SilentlyContinue

$vmAdapters = @()
if ($vms) {
    $vmAdapters = Get-VMNetworkAdapter -VMName $vms.Name -ErrorAction SilentlyContinue
}

$hostAdapters = Get-NetAdapter -Name 'vEthernet*' -ErrorAction SilentlyContinue |
                Where-Object { $_.Status -eq 'Up' }

$vmCount   = ($vmAdapters  | Measure-Object).Count
$hostCount = ($hostAdapters | Measure-Object).Count
$totalSteps = $vmCount + $hostCount

if ($totalSteps -eq 0) {
    Write-Host "No VM or host Hyper-V network adapters found on this host." -ForegroundColor Yellow
    return
}

$step = 0

function Update-ReportProgress {
    param(
        [string]$StatusText
    )
    if ($totalSteps -le 0) { return }
    $script:step++
    $percent = [int](($script:step / $totalSteps) * 100)
    Write-Progress -Activity "Building Hyper-V VLAN report" -Status $StatusText -PercentComplete $percent
}

# ---------------------------
# 1) VM network adapters
# ---------------------------
foreach ($ad in $vmAdapters) {

    Update-ReportProgress -StatusText "Processing VM network adapters ($step of $totalSteps)"

    $v = Get-VMNetworkAdapterVlan -VMName $ad.VMName -VMNetworkAdapterName $ad.Name

    $mode = switch ($v.OperationMode) {
        0 { 'Untagged' }
        1 { 'Access'   }
        2 { 'Trunk'    }
        default { $v.OperationMode }
    }

    $accessVlan   = $null
    $defaultVlan  = $null
    $allowedVlans = $null

    if ($mode -eq 'Access') {
        $accessVlan = $v.AccessVlanId
    }
    elseif ($mode -eq 'Trunk') {
        $defaultVlan = $v.NativeVlanId
        if ($null -ne $v.AllowedVlanIdList) {
            $allowedVlans = $v.AllowedVlanIdList -join ','
        }
    }
    else {
        $defaultVlan = 0
    }

    $ip = $ad.IPAddresses |
        Where-Object { $_ -match '^(?:\d{1,3}\.){3}\d{1,3}$' } |
        Select-Object -First 1

    if (-not $ip -and $DnsSuffix) {
        try {
            $fqdn = "$($ad.VMName).$DnsSuffix"
            $dnsResult = Resolve-DnsName -Name $fqdn -ErrorAction Stop |
                         Where-Object { $_.Type -eq 'A' } |
                         Select-Object -First 1
            if ($dnsResult) {
                $ip = $dnsResult.IPAddress
            }
        }
        catch { }
    }

    $sw = Get-VMSwitch -Name $ad.SwitchName -ErrorAction SilentlyContinue
    $switchType = if ($sw) { $sw.SwitchType } else { '' }

    $rows += [PSCustomObject]@{
        ObjectType   = 'VM'
        Name         = $ad.VMName
        SwitchType   = $switchType
        IPAddress    = $ip
        SwitchName   = $ad.SwitchName
        AdapterName  = $ad.Name
        Mode         = $mode
        AccessVlan   = $accessVlan
        AllowedVlans = $allowedVlans
        DefaultVlan  = $defaultVlan
    }
}

# ---------------------------
# 2) Host vEthernet adapters
# ---------------------------
$hostName = $env:COMPUTERNAME

foreach ($ha in $hostAdapters) {

    Update-ReportProgress -StatusText "Processing host vEthernet adapters ($step of $totalSteps)"

    $ipObj = Get-NetIPAddress -InterfaceIndex $ha.ifIndex -AddressFamily IPv4 -ErrorAction SilentlyContinue |
             Where-Object { $_.IPAddress -notmatch '^169\.254\.' } |
             Select-Object -First 1

    $ip = $ipObj.IPAddress

    $switchName = $ha.Name
    if ($switchName -match '^vEthernet \((.+)\)$') {
        $switchName = $Matches[1]
    }

    $sw = Get-VMSwitch -Name $switchName -ErrorAction SilentlyContinue
    $switchType = if ($sw) { $sw.SwitchType } else { 'Unknown' }

    $rows += [PSCustomObject]@{
        ObjectType   = 'Host'
        Name         = $hostName
        SwitchType   = $switchType
        IPAddress    = $ip
        SwitchName   = $switchName
        AdapterName  = ''
        Mode         = ''
        AccessVlan   = $null
        AllowedVlans = $null
        DefaultVlan  = $null
    }
}

Write-Progress -Activity "Building Hyper-V VLAN report" -Completed

# ---------------------------
# 3) Column selection
# ---------------------------
if ($Compact) {
    $columns = @(
        'ObjectType',
        'Name',
        'SwitchType',
        'IPAddress',
        'SwitchName',
        'Mode',
        'AccessVlan',
        'AllowedVlans',
        'DefaultVlan'
    )
}
else {
    $columns = @(
        'ObjectType',
        'Name',
        'SwitchType',
        'IPAddress',
        'SwitchName',
        'AdapterName',
        'Mode',
        'AccessVlan',
        'AllowedVlans',
        'DefaultVlan'
    )
}

$rows = $rows | Sort-Object `
    @{ Expression = { if ($_.ObjectType -eq 'Host') { 0 } else { 1 } } }, `
    @{ Expression = {
            if ($_.ObjectType -eq 'VM' -and $_.AccessVlan) {
                [int]$_.AccessVlan
            } else {
                [int]::MaxValue
            }
        }
    }, `
    Name, `
    AdapterName

# ---------------------------
# 4) Column widths
# ---------------------------
$widths = @{}
foreach ($col in $columns) {
    $headerLen = $col.Length
    $maxData = ($rows | ForEach-Object {
        $val = [string]($_.$col)
        if ($null -eq $val) { $val = '' }
        $val.Length
    } | Measure-Object -Maximum).Maximum

    if (-not $maxData) { $maxData = 0 }
    $widths[$col] = [Math]::Max($headerLen, $maxData)
}

function Write-Border {
    param($Columns, $Widths)
    $line = '+'
    foreach ($c in $Columns) {
        $line += ('-' * ($Widths[$c] + 2)) + '+'
    }
    Write-Host $line
}

# ---------------------------
# 5) Color helpers
# ---------------------------
function Get-NetworkColor {
    param(
        [string]$AdapterName,
        [string]$SwitchName,
        [string]$ColumnName
    )

    if ($ColumnName -eq 'AdapterName') { return 'White' }

    switch ($SwitchName) {
        'Default Switch' { return 'DarkYellow' }
        'LAN-SW'         { return 'Green' }
        'MGMT-SW'        { return 'Cyan' }
        'WAN-SW'         { return 'Red' }
        default          { return 'White' }
    }
}

function Get-CellColor {
    param($ColumnName, $Row)

    switch ($ColumnName) {
        'ObjectType' {
            if ($Row.ObjectType -eq 'Host') { return 'Blue' }
            else { return 'White' }
        }
        'IPAddress' {
            if ($Row.IPAddress) { return 'White' }
            return 'DarkGray'
        }
        'SwitchType' {
            switch ($Row.SwitchType) {
                'External' { return 'Red' }
                'Internal' { return 'Gray' }
                'Private'  { return 'DarkYellow' }
                'Unknown'  { return 'DarkGray' }
                default    { return 'White' }
            }
        }
        'AdapterName' {
            return Get-NetworkColor -AdapterName $Row.AdapterName -SwitchName $Row.SwitchName -ColumnName 'AdapterName'
        }
        'SwitchName' {
            return Get-NetworkColor -AdapterName $Row.AdapterName -SwitchName $Row.SwitchName -ColumnName 'SwitchName'
        }
        default { return 'White' }
    }
}

# ---------------------------
# 6) HEADER (renamed labels)
# ---------------------------
$columnLabels = @{
    'SwitchType'   = 'Type'
    'SwitchName'   = 'Switch'
    'AdapterName'  = 'Adapter'
    'AccessVlan'   = 'Access'
    'AllowedVlans' = 'Allowed'
    'DefaultVlan'  = 'Default'
}

Write-Border -Columns $columns -Widths $widths

Write-Host -NoNewline '|'
foreach ($c in $columns) {
    $label = if ($columnLabels.ContainsKey($c)) { $columnLabels[$c] } else { $c }
    $text = $label.PadRight($widths[$c])
    Write-Host (" {0} " -f $text) -NoNewline -ForegroundColor Cyan
    Write-Host -NoNewline '|'
}
Write-Host
Write-Border -Columns $columns -Widths $widths

# ---------------------------
# 7) DATA ROWS
# ---------------------------
$lastName = $null

foreach ($row in $rows) {

    if ($lastName -and $lastName -ne $row.Name) {
        Write-Border -Columns $columns -Widths $widths
    }
    $lastName = $row.Name

    Write-Host -NoNewline '|'
    foreach ($c in $columns) {
        $val = [string]$row.$c
        if ($null -eq $val) { $val = '' }

        $text  = $val.PadRight($widths[$c])
        $color = Get-CellColor -ColumnName $c -Row $row

        Write-Host (" {0} " -f $text) -NoNewline -ForegroundColor $color
        Write-Host -NoNewline '|'
    }
    Write-Host
}

Write-Border -Columns $columns -Widths $widths