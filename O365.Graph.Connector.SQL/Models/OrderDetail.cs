using Microsoft.Graph.Models.ExternalConnectors;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Text;
using System.Text.Json.Serialization;
using System.Threading.Tasks;

namespace O365.Graph.Connector.SQL.Models
{
    public class OrderDetail
    {

        public string OrderID { get; set; }
        public string OrderDate { get; set; }
        public string ShipAddress { get; set; }
        public string ShipCountry { get; set; }
        public string ProductID { get; set; }
        public string ProductName { get; set; }
        public string Quantity { get; set; }
        public string UnitPrice { get; set; }
        public string CustomerID { get; set; }
        public string ContactName { get; set; }
        public string CompanyName { get; set; }
        
    }

}
