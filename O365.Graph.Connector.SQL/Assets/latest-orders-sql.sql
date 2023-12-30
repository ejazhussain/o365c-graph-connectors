/*Identify the top 100 latest orders including information about the product, customer name, etc.*/
SELECT
	Top 100
    Orders.OrderID,
    Orders.OrderDate,
	Orders.ShipAddress,
	Orders.ShipCountry,	    
    [dbo].[Order Details].ProductID,
    Products.ProductName,
    [dbo].[Order Details].Quantity,
    [dbo].[Order Details].UnitPrice,
	[dbo].Customers.CustomerId,
    [dbo].Customers.ContactName,
	[dbo].Customers.CompanyName	
FROM [dbo].Orders
INNER
 
JOIN [dbo].[Order Details] ON [dbo].Orders.OrderID = [dbo].[Order Details].OrderID
INNER
 
JOIN [dbo].Customers ON [dbo].Orders.CustomerID = [dbo].Customers.CustomerID
INNER
 
JOIN Products ON [dbo].[Order Details].ProductID = Products.ProductID
ORDER BY Orders.OrderDate DESC