## ---------- Working with SQL Server ---------- ##

## - Get SQL Server Table data:
$SQLServer = 'AP_HOME_PC\SQLEXPRESS';
$Database = 'AdventureWorks2012';
$SqlQuery = @'
select top 3
		p.FirstName
		, p.LastName
		, c.AccountNumber
        , convert(varchar(10), c.ModifiedDate, 101) as ModifiedDate
		, soh.PurchaseOrderNumber
		, convert(varchar(10), soh.OrderDate, 101) as OrderDate
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
	where soh.OrderDate between '2008-06-01' and '2008-12-01'
		and sod.UnitPrice > 2400
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

[PsObject[]]$DStbl = @()
$DStbl += [PsObject]@{ AccountNumber = "AW00000000"; LastName = "BOF"}
$DStbl = $DataSetTable | Select-Object -property AccountNumber, LastName

$Dttxt = (Get-date -UFormat "%A  %m/%d/%Y").ToString()
## ---------- Working with Excel ---------- ##
rm .\output\AW_ExcelExport_Results.xlsx -ErrorAction Ignore
$eFmt = New-ConditionalText -Range "E:E" -
$eTxt = New-ConditionalText AW00011001 blue
$DataSetTable | Select-Object -Property FirstName, LastName, AccountNumber, PurchaseOrderNumber, OrderDate, ModifiedDate | Export-Excel .\output\AW_ExcelExport_Results.xlsx -AutoSize -AutoFilter -title "Date: $Dttxt" -TitleSize 14 -NoNumberConversion OrderDate, ModifiedDate -ConditionalText $eTxt
#$DataSetTable | Select-Object -ExcludeProperty RowError, RowState, Table, ItemArray, HasErrors | Export-Excel .\output\AW_ExcelExport_Results.xlsx -WorksheetName 'sheet2' -AutoFilter
#$excel = $DataSetTable | Select-Object -ExcludeProperty RowError, RowState, Table, ItemArray, HasErrors | Export-Excel .\output\AW_ExcelExport_Results.xlsx -WorksheetName 'sheet2' -ExcludeProperty RowError, RowState, Table, ItemArray, HasErrors -AutoFilter -PassThru
#    $excel.Workbook.Worksheets["Sheet2"].Column(5).style.font.bold = $true
#    $excel.Workbook.Worksheets["Sheet2"].Column(6).style.font.bold = $true 
    $excel.Workbook.Worksheets["Sheet2"].Row(1).style.font.bold = $true 
#    $excel.Save()
#    $excel.Dispose()
    Export-excel $excel -WorksheetName "Sheet2" -IncludePivotChart -ChartType ColumnClustered -NoLegend -PivotRows OrderDate  -PivotData @{'LineTotal'='Sum'}
    Start-Process .\output\AW_ExcelExport_Results.xlsx