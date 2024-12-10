# Import the ImportExcel module
# If not installed, install it using `Install-Module -Name ImportExcel`

# Get the directory of the current script
$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path

# Path to the Excel file in the same directory as the script
$excelPath = Join-Path -Path $scriptPath -ChildPath "raffle.xlsx"

# Check if the Excel file exists
if (-not (Test-Path $excelPath)) {
    Write-Host "Error: The file raffle.xlsx does not exist in the script's directory." -ForegroundColor Red
    exit
}

# Import the data from the Excel file
$data = Import-Excel -Path $excelPath

# Create a list to hold all tickets
$allTickets = @()

# Loop through the rows of the Excel file
foreach ($row in $data) {
    $name = $row.Name          # Column 1: Name
    $ticketCount = $row.Tickets # Column 2: Number of tickets purchased (now "Tickets")

    # Skip rows with missing or invalid data without warning
    if (-not $name -or -not $ticketCount -or $ticketCount -le 0) {
        continue
    }

    # Add the name to the ticket list as many times as the number of tickets
    for ($i = 1; $i -le $ticketCount; $i++) {
        $allTickets += $name
    }
}

# Draw a random winner from the list
if ($allTickets.Count -eq 0) {
    Write-Host "Error: No tickets available for the draw. Please check the Excel file." -ForegroundColor Red
    exit
}

$winner = $allTickets | Get-Random

# Get the current date and time in German format
$currentDateTime = Get-Date -Format "dd.MM.yyyy HH:mm"

# Create the results data with English column names
$results = @(
    [PSCustomObject]@{
        DateTime = $currentDateTime
        Winner   = $winner
    }
)

# Append or create results in the Excel file
if (Test-Path $excelPath) {
    $results | Export-Excel -Path $excelPath -WorksheetName "Results" -Append
} else {
    $results | Export-Excel -Path $excelPath -WorksheetName "Results"
}

# Release potential locks on the file by closing it properly
Start-Sleep -Seconds 1

# Output the winner and draw time in German format to the console
Write-Host "Winner: $winner"
Write-Host "Winner drawn at $currentDateTime"
