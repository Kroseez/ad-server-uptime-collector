# Script parameters
param (
    [string]$OutputDirectory = "C:\Logs\Uptime",
    [string]$OutputFileName = "Uptime_Report_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
)

# Full path to the report file
$OutputFile = Join-Path -Path $OutputDirectory -ChildPath $OutputFileName

# Check and create a directory if necessary
if (-not (Test-Path -Path $OutputDirectory)) {
    try {
        New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null
        Write-Host "Directory created $OutputDirectory" -ForegroundColor Green
    }
    catch {
        Write-Host "Error creating directory: $_" -ForegroundColor Red
        exit 1
    }
}

# Importing the ActiveDirectory module
try {
    Import-Module ActiveDirectory -ErrorAction Stop
}
catch {
    Write-Host "ActiveDirectory module not found. Make sure RSAT is installed." -ForegroundColor Red
    exit 1
}

# We get a list of all servers from OU=Servers and all nested OUs
try {
    Write-Host "Search for computers in AD..." -ForegroundColor Cyan
    $servers = Get-ADComputer -Filter * -SearchBase "OU=Servers,DC=domain,DC=com" -SearchScope Subtree -Properties Name, LastLogonDate | # Enter the address of your OU where the servers are stored
               Select-Object -ExpandProperty Name
    Write-Host "Servers found: $($servers.Count)" -ForegroundColor Cyan
}
catch {
    Write-Host "Error searching for computers in AD: $_" -ForegroundColor Red
    exit 1
}

# Initialize an array for the results
$results = @()

# We process each server
foreach ($server in $servers) {
    Write-Host "Checking the server: $server" -ForegroundColor Gray
    
    # Checking server availability
    $online = Test-Connection -ComputerName $server -Count 1 -Quiet -ErrorAction SilentlyContinue
    
    if ($online) {
        try {
            # Getting information about the system
            $os = Get-WmiObject -Class Win32_OperatingSystem -ComputerName $server -ErrorAction Stop
            
            # Calculating uptime
            $lastBootTime = $os.ConvertToDateTime($os.LastBootUpTime)
            $uptime = (Get-Date) - $lastBootTime
            $uptimeString = "{0} d. {1} h. {2} min." -f $uptime.Days, $uptime.Hours, $uptime.Minutes
            
            $status = "Online"
            Write-Host "$server : Online (Uptime: $uptimeString)" -ForegroundColor Green
        }
        catch {
            $status = "WMI Error"
            $uptimeString = "N/A"
            $lastBootTime = "N/A"
            Write-Host "$server : Online (WMI Error: $($_.Exception.Message))" -ForegroundColor Yellow
        }
    }
    else {
        $status = "Offline"
        $uptimeString = "N/A"
        $lastBootTime = "N/A"
        Write-Host "$server : Offline" -ForegroundColor Red
    }
    
    # Add the result to the array
    $results += [PSCustomObject]@{
        ServerName    = $server
        Status        = $status
        LastBootTime  = $lastBootTime
        Uptime        = $uptimeString
        CheckDate     = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
    }
}

# Export results to CSV
try {
    $results | Export-Csv -Path $OutputFile -NoTypeInformation -Encoding UTF8 -Delimiter ";"
    Write-Host "The report was successfully saved: $OutputFile" -ForegroundColor Green
    
    # Open the folder with the report
    Invoke-Item -Path $OutputDirectory
}
catch {
    Write-Host "Error saving report: $_" -ForegroundColor Red
}