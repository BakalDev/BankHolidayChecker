# Variables
$today = (Get-Date -Format dd/MM/yyyy)
$todayIsBankHoliday = $false # Updated from IsBankHoliday function
$date = Get-Date # Used for determining the day of the week
$dayOfWeek = [int]$date.dayofWeek # Returns int of day of week (Sunday = 0)

# Check if today is bank holiday
IsBankHoliday $today

# So we know today is a bank holiday
# Now what? Do we want to check if it's a monday?
if ($todayIsBankHoliday -eq $true -AND $dayOfWeek -eq 1) {
Write-Host "Ignore, this is fine"
}

function IsBankHoliday([string]$arg1) {

# Public Holidays Variables
$publicHolidaysFolderLocation = "C:\Users\adambd\Desktop\"
$publicHolidaysFileName = "Public Holidays.xlsx"
$publicHolidaysSheetName = "Current"

# Load excel object
$objExcel = New-Object -ComObject Excel.Application
$workbook = $objExcel.Workbooks.Open($publicHolidaysFolderLocation + $publicHolidaysFileName)
$sheet = $workbook.Worksheets.Item($publicHolidaysSheetName)
$objExcel.Visible = $false

# Check every value on every column
$rowMax = ($sheet.UsedRange.Rows).count

# Set starting positions
$rowDate, $colDate = 1, 1

# Loop through each row and store each variable
for ($i = 1; $i -le $rowMax-1; $i++) {

$date = $sheet.Cells.Item($rowDate+$i, $colDate).text
	
	if ($date -eq $arg1) {
		$global:todayIsBankHoliday = $true
	}
}

# quit excel
$objExcel.quit()

}
