## ---------- Working with SQL Server ---------- ##

## - Get SQL Server Table data:
$SQLServer = 'AP_HOME_PC\SQLEXPRESS';
$Database = 'AdventureWorks2012';
$SqlQuery = @'
select 
		p.FirstName
		, p.LastName
		, c.AccountNumber
		, soh.PurchaseOrderNumber
		, soh.OrderDate
		, sod.ProductID
		, sod.OrderQty
		, sod.UnitPrice
		, sod.UnitPriceDiscount
		, sod.LineTotal
from 
	[Sales].[Customer] c
	inner join [Person].[Person] p on c.PersonID = p.BusinessEntityID
	inner join [Sales].[SalesOrderHeader] soh on c.CustomerID = soh.CustomerID
	inner join [Sales].[SalesOrderDetail] sod on soh.SalesOrderID = sod.SalesOrderID
	where soh.OrderDate between '2008-06-01' and '2008-06-01'
	--	and sod.UnitPrice > 2400
'@;

## - Connect to SQL Server using non-SMO class 'System.Data':
$SqlConnection = New-Object System.Data.SqlClient.SqlConnection;
$SqlConnection.ConnectionString = "Server = $SQLServer; Database = $Database; Integrated Security = True";

$SqlCmd = New-Object System.Data.SqlClient.SqlCommand;
$SqlCmd.CommandText = $SqlQuery;
$SqlCmd.Connection = $SqlConnection;

## - Extract and build the SQL data object '$DataSetTable':
$SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter;
$SqlAdapter.SelectCommand = $SqlCmd;
$DataSet = New-Object System.Data.DataSet;
$SqlAdapter.Fill($DataSet);
$DataSetTable = $DataSet.Tables["Table"];

## ---------- Working with Excel File ---------- ##

## - Create an Excel Application instance:
$xlsObj = New-Object -ComObject Excel.Application;

## - Create new Workbook and Sheet (Visible = 1 / 0 not visible)
$xlsObj.Visible = 0;
$xlsWb = $xlsobj.Workbooks.Add();
$xlsSh = $xlsWb.Worksheets.item(1);

## - Build the Excel column heading:
[Array] $getColumnNames = $DataSetTable.Columns | Select ColumnName;

## - Build column header:
[Int] $RowHeader = 1;
foreach ($ColH in $getColumnNames)
{
$xlsSh.Cells.item(1, $RowHeader).font.bold = $true;
$xlsSh.Cells.item(1, $RowHeader).font.underline = $true;
$xlsSh.Cells.item(1, $RowHeader).Font.Name = "Cambria";
$xlsSh.Cells.item(1, $RowHeader).Font.Size = 14;
$xlsSh.Cells.item(1, $RowHeader).Font.Color = 8210719;
$xlsSh.Cells.item(1, $RowHeader) = $ColH.ColumnName;
$RowHeader++;
};

## - Adding the data start in row 2 column 1:
[Int] $rowData = 2;
[Int] $colData = 1;

foreach ($rec in $DataSetTable.Rows)
{
foreach ($Coln in $getColumnNames)
{
## - Next line convert cell to be text only:
$xlsSh.Cells.NumberFormat = "@";

## - Populating columns:
$xlsSh.Cells.Item($rowData, $colData) = `
$rec.$($Coln.ColumnName).ToString();
$ColData++;
};
$rowData++; $ColData = 1;
};

## - Adjusting columns in the Excel sheet:
$xlsRng = $xlsSH.usedRange;
$xlsRng.EntireColumn.AutoFit();

## ---------- Saving file and Terminating Excel Application ---------- ##

## - Saving Excel file - if the file exist do delete then save
$xlsFile = "C:\Users\Arturo\Documents\Powershell\test\AW_201106ExceldbResults_$((Get-Date).ToString("yyyyMMdd")).xlsx";

if (Test-Path $xlsFile)
{
Remove-Item $xlsFile
$xlsObj.ActiveWorkbook.SaveAs($xlsFile);
}
else
{
$xlsObj.ActiveWorkbook.SaveAs($xlsFile);
};
$xlsObj.ActiveWorkbook.Close($xlsFile);

## Quit Excel and Terminate Excel Application process:
$xlsObj.Quit(); #(Get-Process Excel*) | foreach ($_) { $_.kill() };
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($xlsObj) | Out-Null;


## - End of Script - ##
