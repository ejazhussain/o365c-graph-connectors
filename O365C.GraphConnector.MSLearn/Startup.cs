using Microsoft.Azure.Functions.Extensions.DependencyInjection;
using Microsoft.Azure.WebJobs.Host.Bindings;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using O365C.GraphConnector.MSLearn.Models;
using O365C.GraphConnector.MSLearn.Services;
using System.Text.Json;
using O365C.GraphConnector.MSLearn;

[assembly: FunctionsStartup(typeof(Startup))]
namespace O365C.GraphConnector.MSLearn
{
    public class Startup : FunctionsStartup
    {
        public override void Configure(IFunctionsHostBuilder builder)
        {

            // Get the configuration from the app settings
            var config = builder.GetContext().Configuration;
            var azureFunctionSettings = new AzureFunctionSettings();
            config.Bind(azureFunctionSettings);
            // Add our configuration class
            builder.Services.AddSingleton(options => { return azureFunctionSettings; });

            // Register your services
            builder.Services.AddSingleton<IAccessTokenProvider, AccessTokenProvider>();
            builder.Services.AddHttpClient();
            builder.Services.AddSingleton<IGraphService, GraphService>();         
            builder.Services.AddSingleton<ILearnCatalogService, LearnCatalogService>();


            // Configure the JSON serializer to use camelCase property names
            builder.Services.Configure<JsonSerializerOptions>(options =>
            {
                options.PropertyNamingPolicy = JsonNamingPolicy.CamelCase;
            });

            // Register the ExecutionContextOptions
            builder.Services.AddSingleton<ExecutionContextOptions>();

        }
    }
}
