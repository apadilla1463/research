select 
		p.FirstName
		, p.LastName
		, c.AccountNumber
--		, soh.PurchaseOrderNumber
		, soh.OrderDate
--		, sod.ProductID
--		, sod.OrderQty
--		, sod.UnitPrice
--		, sod.UnitPriceDiscount
--		, sod.LineTotal
from 
	[Sales].[Customer] c
	inner join AdventureWorks2012.[Person].[Person] p on c.PersonID = p.BusinessEntityID
	inner join AdventureWorks2012.[Sales].[SalesOrderHeader] soh on c.CustomerID = soh.CustomerID
--	inner join AdventureWorks2012.[Sales].[SalesOrderDetail] sod on soh.SalesOrderID = sod.SalesOrderID
	where soh.OrderDate between '2008-06-01' and '2008-06-10'
		--and sod.UnitPrice > 2400