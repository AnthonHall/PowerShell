<#Usage
-all Exports all data
-startDate Exports data after this date
-endDate Exports data before this date
#>

param (
    [switch] $all,
    [datetime] $startDate,
    [datetime] $endDate
)

#Windows Authentication
$connectstring = "DSN=QuickBooks POS Data;OLE DB Services=-2;"
$sql = "select * from SalesReceipt"

$conn = New-Object System.Data.Odbc.OdbcConnection($connectstring)
$conn.Open()
$cmd = New-Object system.Data.Odbc.OdbcCommand($sql,$conn)
$da = New-Object system.Data.Odbc.OdbcDataAdapter($cmd)
$dt = New-Object system.Data.datatable
$null = $da.fill($dt)
$conn.close()
$date = Get-Date -UFormat "%m%d%Y"


If ($all) {
    #All data
    $filename = "\\overstock.com\REDACTED\SalesReceipt_" + $date + "_All.csv"
    $dt | Export-Csv $filename -notype
    }
Else {
    If (!$startDate -and !$endDate) {
        #Yesterday's data
        $filename = "\\overstock.com\REDACTED\SalesReceipt_" + $date + ".csv"
        $dt | Where TxnDate -gt (Get-Date).AddDays(-2) | Export-Csv $filename -notype
        }
    Else {
        #Specified date range
        If (!$startDate) {$startDate = (Get-Date -month 01 -day 01 -year 1970)}
        If (!$endDate) {$endDate = (Get-Date).AddDays(1)}
        If ($startDate -gt $endDate) {
            $tempDate = $startDate
            $startDate = $endDate
            $endDate = $tempDate
            }
        $startDateDOS = Get-Date $startDate -UFormat "%m%d%Y"
        $endDateDOS = Get-Date $endDate -UFormat "%m%d%Y"
        $filename = "\\overstock.com\REDACTED\SalesReceipt_" + $startDateDOS + "-" + $endDateDOS + ".csv"
        Write-Host $filename
        $dt | where TxnDate -gt $startDate | Where TxnDate -lt $endDate | Export-Csv $filename -notype
        }
    }
