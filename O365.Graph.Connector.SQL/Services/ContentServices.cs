using Dapper;
using Microsoft.Graph.Models.ExternalConnectors;
using Newtonsoft.Json;
using O365.Graph.Connector.SQL.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace O365.Graph.Connector.SQL.Services
{
    public class ContentServices
    {
        public static IEnumerable<OrderDetail> Extract(string ConnectionString)
        {
            List<OrderDetail> orders = new List<OrderDetail>();

            try
            {

                using (IDbConnection db = new SqlConnection(ConnectionString))
                {
                    string sqlQuery = "SELECT\r\n\tTop 100\r\n    Orders.OrderID,\r\n    Orders.OrderDate,\r\n\tOrders.ShipAddress,\r\n\tOrders.ShipCountry,\t    \r\n    [dbo].[Order Details].ProductID,\r\n    Products.ProductName,\r\n    [dbo].[Order Details].Quantity,\r\n    [dbo].[Order Details].UnitPrice,\r\n\t[dbo].Customers.CustomerId,\r\n    [dbo].Customers.ContactName,\r\n\t[dbo].Customers.CompanyName\t\r\nFROM [dbo].Orders\r\nINNER\r\n \r\nJOIN [dbo].[Order Details] ON [dbo].Orders.OrderID = [dbo].[Order Details].OrderID\r\nINNER\r\n \r\nJOIN [dbo].Customers ON [dbo].Orders.CustomerID = [dbo].Customers.CustomerID\r\nINNER\r\n \r\nJOIN Products ON [dbo].[Order Details].ProductID = Products.ProductID\r\nORDER BY Orders.OrderDate DESC";
                    orders = db.Query<OrderDetail>(sqlQuery).ToList();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            return orders;
        }


        public static IEnumerable<ExternalItem> Transform(IEnumerable<OrderDetail> content)
        {

            return content.Select(a =>
            {
                return new ExternalItem
                {      
                    Id = a.OrderID,
                    Acl = new()
                    {
                            new()
                            {
                                Type = AclType.Everyone,
                                Value = "everyone",
                                AccessType = AccessType.Grant,
                            },
                    },
                    Properties = new()
                    {
                        AdditionalData = new Dictionary<string, object> {

                                { "OrderID", a.OrderID ?? "" },
                                { "OrderDate", a.OrderDate ?? "" },
                                { "ShipAddress", a.ShipAddress },
                                { "ShipCountry" , a.ShipCountry },
                                { "ProductID", a.ProductID ?? "" },
                                { "ProductName", a.ProductName ?? "" },
                                { "Quantity", a.Quantity ?? "" },
                                { "UnitPrice", a.UnitPrice ?? "" },
                                { "CustomerID", a.CustomerID ?? "" },
                                { "ContactName", a.ContactName ?? "" },
                                { "CompanyName", a.CompanyName ?? "" },
                        }
                    },
                    Activities = new()
            {
                new()
                {
                    OdataType = "#microsoft.graph.externalConnectors.externalActivity",
                    Type = ExternalActivityType.Created,
                    StartDateTime = new DateTime(),
                    PerformedBy = new Identity
                    {
                          Type = IdentityType.User,
                          Id = "2a5de346-1d63-4c7a-897f-b1f4b5316fe5"
                    }

                },
            },
                    Content = new()
                    {
                        Value = $"{a.OrderID},{a.OrderDate},{a.ProductName}" ?? "",
                        Type = ExternalItemContentType.Text
                    },


                };
            });
        }

    }
}
