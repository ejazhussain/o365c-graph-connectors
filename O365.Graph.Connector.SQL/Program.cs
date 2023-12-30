// See https://aka.ms/new-console-template for more information
using Microsoft.Graph;
using Microsoft.Graph.Models.ExternalConnectors;
using Microsoft.Graph.Models.ODataErrors;
using O365.Graph.Connector.SQL;
using O365.Graph.Connector.SQL.Services;
using System.Text.Json;

Console.WriteLine("Hello, World!");
Console.WriteLine("Northwind Inventory Search Connector\n");

var settings = Settings.LoadSettings();

// Initialize Graph
InitializeGraph(settings);

ExternalConnection? currentConnection = null;
int choice = -1;

while (choice != 0)
{
    Console.WriteLine($"Current connection: {(currentConnection == null ? "NONE" : currentConnection.Name)}\n");
    Console.WriteLine("Please choose one of the following options:");
    Console.WriteLine("0. Exit");
    Console.WriteLine("1. Create a connection");
    Console.WriteLine("2. Select an existing connection");
    Console.WriteLine("3. Delete current connection");
    Console.WriteLine("4. Register schema for current connection");
    Console.WriteLine("5. View schema for current connection");
    Console.WriteLine("6. Push updated items to current connection");
    Console.WriteLine("7. Push ALL items to current connection");
    Console.Write("Selection: ");

    try
    {
        choice = int.Parse(Console.ReadLine() ?? string.Empty);
    }
    catch (FormatException)
    {
        // Set to invalid value
        choice = -1;
    }

    switch (choice)
    {
        case 0:
            // Exit the program
            Console.WriteLine("Goodbye...");
            break;
        case 1:
            currentConnection = await CreateConnectionAsync();
            break;
        case 2:
            currentConnection = await SelectExistingConnectionAsync();
            break;
        case 3:
            await DeleteCurrentConnectionAsync(currentConnection);
            currentConnection = null;
            break;
        case 4:
            await RegisterSchemaAsync();
            break;
        case 5:
            await GetSchemaAsync();
            break;
        case 6:
            await UpdateItemsFromDatabaseAsync(false, settings);
            break;
        //case 7:
        //    await UpdateItemsFromDatabaseAsync(false, settings.TenantId);
        //    break;
        default:
            Console.WriteLine("Invalid choice! Please try again.");
            break;
    }

}



    void InitializeGraph(Settings settings)
{
    try
    {
        GraphService.Initialize(settings);
    }
    catch (Exception ex)
    {
        Console.WriteLine($"Error initializing Graph: {ex.Message}");
    }
}

async Task<ExternalConnection?> CreateConnectionAsync()
{
   
    try
    {
        // Create the connection
        var connection = await GraphService.CreateConnectionAsync();
        Console.WriteLine($"New connection created - Name: {connection?.Name}, Id: {connection?.Id}");
        return connection;
    }
    catch (ODataError odataError)
    {
        Console.WriteLine($"Error creating connection: {odataError.ResponseStatusCode}: {odataError.Error?.Code} {odataError.Error?.Message}");
        return null;
    }
}

async Task<ExternalConnection?> SelectExistingConnectionAsync()
{
    // TODO
    Console.WriteLine("Getting existing connections...");
    try
    {
        var response = await GraphService.GetExistingConnectionsAsync();
        var connections = response?.Value ?? new List<ExternalConnection>();
        if (connections.Count <= 0)
        {
            Console.WriteLine("No connections exist. Please create a new connection");
            return null;
        }

        // Display connections
        Console.WriteLine("Choose one of the following connections:");
        var menuNumber = 1;
        foreach (var connection in connections)
        {
            Console.WriteLine($"{menuNumber++}. {connection.Name}");
        }

        ExternalConnection? selection = null;

        do
        {
            try
            {
                Console.Write("Selection: ");
                var choice = int.Parse(Console.ReadLine() ?? string.Empty);
                if (choice > 0 && choice <= connections.Count)
                {
                    selection = connections[choice - 1];
                }
                else
                {
                    Console.WriteLine("Invalid choice.");
                }
            }
            catch (FormatException)
            {
                Console.WriteLine("Invalid choice.");
            }
        } while (selection == null);

        return selection;
    }
    catch (ODataError odataError)
    {
        Console.WriteLine($"Error getting connections: {odataError.ResponseStatusCode}: {odataError.Error?.Code} {odataError.Error?.Message}");
        return null;
    }
}

async Task DeleteCurrentConnectionAsync(ExternalConnection? connection)
{
    if (connection == null)
    {
        Console.WriteLine(
            "No connection selected. Please create a new connection or select an existing connection.");
        return;
    }

    try
    {
        await GraphService.DeleteConnectionAsync(connection.Id);
        Console.WriteLine($"{connection.Name} deleted successfully.");
    }
    catch (ODataError odataError)
    {
        Console.WriteLine($"Error deleting connection: {odataError.ResponseStatusCode}: {odataError.Error?.Code} {odataError.Error?.Message}");
    }
}

async Task RegisterSchemaAsync()
{
    if (currentConnection == null)
    {
        Console.WriteLine("No connection selected. Please create a new connection or select an existing connection.");
        return;
    }

    Console.WriteLine("Registering schema, this may take a moment...");

    try
    {
        // Create the schema        
        await GraphService.RegisterSchemaAsync(currentConnection.Id, ConnectionConfiguration.Schema);
        Console.WriteLine("Schema registered successfully");
    }
    catch (ServiceException serviceException)
    {
        Console.WriteLine($"Error registering schema: {serviceException.ResponseStatusCode} {serviceException.Message}");
    }
    catch (ODataError odataError)
    {
        Console.WriteLine($"Error registering schema: {odataError.ResponseStatusCode}: {odataError.Error?.Code} {odataError.Error?.Message}");
    }
}

async Task GetSchemaAsync()
{
    if (currentConnection == null)
    {
        Console.WriteLine("No connection selected. Please create a new connection or select an existing connection.");
        return;
    }

    try
    {
        var schema = await GraphService.GetSchemaAsync(currentConnection.Id);
        Console.WriteLine(JsonSerializer.Serialize(schema));

    }
    catch (ODataError odataError)
    {
        Console.WriteLine($"Error getting schema: {odataError.ResponseStatusCode}: {odataError.Error?.Code} {odataError.Error?.Message}");
    }
}

async Task UpdateItemsFromDatabaseAsync(bool uploadModifiedOnly, Settings settings)
{
    if (currentConnection == null)
    {
        Console.WriteLine("No connection selected. Please create a new connection or select an existing connection.");
        return;
    }

    var content = ContentServices.Extract(settings.ConnectionString);
    var transformedItems = ContentServices.Transform(content);

    foreach (var item in transformedItems)
    {
        Console.Write(string.Format("Loading item {0}...", item.Id));
        try
        {
            await GraphService.AddOrUpdateItemAsync(currentConnection.Id, item);            
            Console.WriteLine("DONE");
        }
        catch (Exception ex)
        {
            Console.WriteLine("ERROR");
            Console.WriteLine(ex.Message);
        }
    }  
}