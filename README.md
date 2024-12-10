
# 🎟️ Raffle Winner Draw Script

This PowerShell script randomly selects a winner from a pool of participants based on ticket entries listed in an Excel spreadsheet. It outputs the winner's name and draw timestamp to the console and logs the results to an Excel file.

---

## 🛠️ Prerequisites

Before running this script, ensure you have:

1. **PowerShell Module:**  
   Install the `ImportExcel` module if it's not already installed:
   ```powershell
   Install-Module -Name ImportExcel
   ```

2. **Excel Data File:**  
   The script expects the input data to be in a file named `raffle.xlsx`. Ensure this Excel file exists and has the following structure:

   | **Name**    | **Lose** |
   |-------------|---------|
   | Alice       | 5       |
   | Bob         | 3       |
   | Charlie     | 2       |

   - **Column 1:** `Name` – Participant's name.  
   - **Column 2:** `Lose` – The number of tickets each participant has purchased.

3. **File Path:**  
   The script expects this file to exist at:  
   ```
   C:\Users\yourusername\Desktop\Teilnehmerliste.xlsx
   ```
   Update the `$excelPath` variable in the script if your file is located elsewhere.

---

## 🚀 What the Script Does

### 1. **Reads Data from the Excel File**
   - Imports participants and their ticket counts from the Excel file.

### 2. **Generates a List of Tickets**
   - Each participant's name is added to a virtual "ticket pool" based on the number of tickets they purchased.  
     Example:  
     If Alice purchased **5 tickets**, her name will be added 5 times to the ticket pool.

### 3. **Randomly Draws a Winner**
   - Randomly selects one name from the "ticket pool."

### 4. **Logs Results to Excel**
   - The draw result (date and winner) is logged to a worksheet named `Results` within the Excel file.
   - If the worksheet already exists, new results will be appended.
   - If the Excel file does not exist, it will be created.

### 5. **Outputs to Console**
   - The winner's name and the draw timestamp are displayed in German date format:
   ```
   Winner: Alice
   Winner drawn at 10.12.2024 09:22
   ```

---

## ⚙️ How to Use

1. **Setup Prerequisites:**  
   Ensure `ImportExcel` is installed and the input file is ready with proper data. Modify `$excelPath` if necessary.

2. **Run the Script:**  
   Run the script using PowerShell:
   ```powershell
   .\YourScriptName.ps1
   ```

3. **Check Results:**  
   - Results will be saved to the `Results` worksheet in the input Excel file.
   - The console will show:
   ```
   Winner: [Winner's Name]
   Winner drawn at [Date and Time in German format]
   ```

---

## 📊 Results

The draw result is saved to the `Results` worksheet in the Excel file, with the following columns:

| **DateTime**         | **Winner** |
|----------------------|------------|
| 10.12.2024 09:22      | Alice      |

---

## 🐛 Troubleshooting

### 1. **File Write Errors**
If you encounter issues writing to the Excel file:
- Ensure you have write permissions to the file's directory.
- Close Excel if it is open to avoid locking the file.

### 2. **Missing `ImportExcel` Module**
If the script fails, ensure you install the `ImportExcel` module:
```powershell
Install-Module -Name ImportExcel
```

---

## 📄 License
This script is licensed under [MIT License](https://opensource.org/licenses/MIT).  
Feel free to fork, modify, or share.
