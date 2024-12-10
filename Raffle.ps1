# Import the ImportExcel module
# If not installed, install it using `Install-Module -Name ImportExcel`

# Path to the Excel file
$excelPath = "C:\raffle.xlsx"

# Import the data from the Excel file
$data = Import-Excel -Path $excelPath

# Create a list to hold all tickets
$allTickets = @()

# Loop through the rows of the Excel file
foreach ($row in $data) {
    $name = $row.Name         # Column 1: Name
    $ticketCount = $row.Lose  # Column 2: Number of tickets purchased

    # Add the name to the ticket list as many times as the number of tickets
    for ($i = 1; $i -le $ticketCount; $i++) {
        $allTickets += $name
    }
}

# Draw a random winner from the list
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

# Check if the Excel file exists
if (Test-Path $excelPath) {
    # Append results to the "Results" worksheet
    $results | Export-Excel -Path $excelPath -WorksheetName "Results" -Append
} else {
    # Create a new Excel file with the "Results" worksheet
    $results | Export-Excel -Path $excelPath -WorksheetName "Results"
}

# Release potential locks on the file by closing it properly
Start-Sleep -Seconds 1

# Output the winner and draw time in German format to the console
Write-Host "Winner: $winner"
Write-Host "Winner drawn at $currentDateTime"
