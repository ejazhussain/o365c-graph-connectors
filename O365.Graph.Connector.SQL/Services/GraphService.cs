using Azure.Identity;
using Microsoft.Graph;
using Microsoft.Extensions.Configuration;
using O365.Graph.Connector.SQL;
using Microsoft.Graph.Models.ExternalConnectors;
using Microsoft.Graph.Authentication;

public static class GraphService
{
    private static GraphServiceClient? graphClient;
    private static HttpClient? httpClient;
    public static void Initialize(Settings settings)
    {
        // Create a credential that uses the client credentials
        // authorization flow
        var credential = new ClientSecretCredential(
            settings.TenantId, settings.ClientId, settings.ClientSecret);

        // Create an HTTP client
        httpClient = GraphClientFactory.Create();

        // Create an auth provider
        var authProvider = new AzureIdentityAuthenticationProvider(
            credential, scopes: new[] { "https://graph.microsoft.com/.default" });

        // Create a Graph client using the credential
        graphClient = new GraphServiceClient(httpClient, authProvider);
    }

    public static async Task<ExternalConnection?> CreateConnectionAsync()
    {
        _ = graphClient ?? throw new MemberAccessException("graphClient is null");
        
        return await graphClient.External.Connections.PostAsync(ConnectionConfiguration.ExternalConnection);
    }

    public static async Task<ExternalConnectionCollectionResponse?> GetExistingConnectionsAsync()
    {
        _ = graphClient ?? throw new MemberAccessException("graphClient is null");

        return await graphClient.External.Connections.GetAsync();
    }

    public static async Task DeleteConnectionAsync(string? connectionId)
    {
        _ = graphClient ?? throw new MemberAccessException("graphClient is null");
        _ = connectionId ?? throw new ArgumentException("connectionId is required");

        await graphClient.External.Connections[connectionId].DeleteAsync();
    }

    public static async Task RegisterSchemaAsync(string? connectionId, Schema schema)
    {
        _ = graphClient ?? throw new MemberAccessException("graphClient is null");
        _ = httpClient ?? throw new MemberAccessException("httpClient is null");
        _ = connectionId ?? throw new ArgumentException("connectionId is required");
        // Use the Graph SDK's request builder to generate the request URL
        var requestInfo = graphClient.External
            .Connections[connectionId]
            .Schema
            .ToGetRequestInformation();

        requestInfo.SetContentFromParsable(graphClient.RequestAdapter, "application/json", schema);

        // Convert the SDK request to an HttpRequestMessage
        var requestMessage = await graphClient.RequestAdapter
            .ConvertToNativeRequestAsync<HttpRequestMessage>(requestInfo);
        _ = requestMessage ?? throw new Exception("Could not create native HTTP request");
        requestMessage.Method = HttpMethod.Post;
        requestMessage.Headers.Add("Prefer", "respond-async");

        // Send the request
        var responseMessage = await httpClient.SendAsync(requestMessage) ??
            throw new Exception("No response returned from API");

        if (responseMessage.IsSuccessStatusCode)
        {
            // The operation ID is contained in the Location header returned
            // in the response
            var operationId = responseMessage.Headers.Location?.Segments.Last() ??
                throw new Exception("Could not get operation ID from Location header");
            await WaitForOperationToCompleteAsync(connectionId, operationId);
        }
        else
        {
            throw new ServiceException("Registering schema failed",
                responseMessage.Headers, (int)responseMessage.StatusCode);
        }
    }

    private static async Task WaitForOperationToCompleteAsync(string connectionId, string operationId)
    {
        _ = graphClient ?? throw new MemberAccessException("graphClient is null");

        do
        {
            var operation = await graphClient.External
                .Connections[connectionId]
                .Operations[operationId]
                .GetAsync();

            if (operation?.Status == ConnectionOperationStatus.Completed)
            {
                return;
            }
            else if (operation?.Status == ConnectionOperationStatus.Failed)
            {
                throw new ServiceException($"Schema operation failed: {operation?.Error?.Code} {operation?.Error?.Message}");
            }

            // Wait 5 seconds and check again
            await Task.Delay(5000);
        } while (true);
    }

    public static async Task<Schema?> GetSchemaAsync(string? connectionId)
    {
        _ = graphClient ?? throw new MemberAccessException("graphClient is null");
        _ = connectionId ?? throw new ArgumentException("connectionId is null");

        return await graphClient.External
            .Connections[connectionId]
            .Schema
            .GetAsync();
    }
    public static async Task AddOrUpdateItemAsync(string? connectionId, ExternalItem item)
    {
        _ = graphClient ?? throw new MemberAccessException("graphClient is null");
        _ = connectionId ?? throw new ArgumentException("connectionId is null");

        await graphClient.External
            .Connections[connectionId]
            .Items[item.Id]
            .PutAsync(item);
    }

}