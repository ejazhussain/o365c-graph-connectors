using System.Text.Json;
using Microsoft.Graph.Models;
using Microsoft.Graph.Models.ExternalConnectors;

public static class ConnectionConfiguration
{
    private static Dictionary<string, object>? _layout;
    private static Dictionary<string, object> Layout
    {
        get
        {
            if (_layout is null)
            {
                var adaptiveCard = File.ReadAllText("resultLayout.json");
                _layout = JsonSerializer.Deserialize<Dictionary<string, object>>(adaptiveCard);
            }

            return _layout!;
        }
    }

    public static ExternalConnection ExternalConnection
    {
        get
        {
            return new ExternalConnection
            {
                Id = "o365cnorthwindsqldb",
                Name = "O365C Northwind SQL Database",
                Description = "The Northwind database contains the sales data for a company called “Northwind Traders,” which imports and exports specialty foods from around the world.",
                SearchSettings = new()
                {
                    SearchResultTemplates = new() {
                    new()
                    {
                        Id = "o365cnwqldb",
                        Priority = 1000,
                        Layout = new Json {
                        AdditionalData = Layout
                        }
                    }
                 }
                }
            };
        }
    }

    public static Schema Schema
    {
        get
        {
            return new Schema
            {
                BaseType = "microsoft.graph.externalItem",
                Properties = new()
        {
          new Property
          {
            Name = "OrderID",
            Type = PropertyType.String,
            IsQueryable = true,
            IsSearchable = true,
            IsRetrievable = true
          },
           new Property
          {
            Name = "OrderDate",
            Type = PropertyType.String,
            IsQueryable = true,
            IsSearchable = true,
            IsRetrievable = true
          },
          new Property
          {
            Name = "ShipAddress",
            Type = PropertyType.String,
            IsQueryable = true,
            IsSearchable = true,
            IsRetrievable = true
          },
          new Property
          {
            Name = "ShipCountry",
           Type = PropertyType.String,
            IsQueryable = true,
            IsSearchable = true,
            IsRetrievable = true
          },
          new Property
          {
             Name = "ProductID",
           Type = PropertyType.String,
            IsQueryable = true,
            IsSearchable = true,
            IsRetrievable = true
          },
          new Property
          {
            Name = "ProductName",
           Type = PropertyType.String,
            IsQueryable = true,
            IsSearchable = true,
            IsRetrievable = true
          },
          new Property
          {
             Name = "Quantity",
           Type = PropertyType.String,
            IsQueryable = true,
            IsSearchable = true,
            IsRetrievable = true
          },
           new Property
          {
             Name = "UnitPrice",
           Type = PropertyType.String,
            IsQueryable = true,
            IsSearchable = true,
            IsRetrievable = true
          },
            new Property
          {
             Name = "CustomerId",
           Type = PropertyType.String,
            IsQueryable = true,
            IsSearchable = true,
            IsRetrievable = true
          },
             new Property
          {
             Name = "ContactName",
           Type = PropertyType.String,
            IsQueryable = true,
            IsSearchable = true,
            IsRetrievable = true
          },
              new Property
          {
             Name = "CompanyName",
           Type = PropertyType.String,
            IsQueryable = true,
            IsSearchable = true,
            IsRetrievable = true
          },

        }
            };
        }
    }
}