using System;
using System.IO;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using System.Collections.Generic;
using O365C.GraphConnector.MSLearn.Services;
using O365C.GraphConnector.MSLearn.Models;
using Microsoft.Graph.Models.ExternalConnectors;
using O365C.GraphConnector.MSLearn.Utils;

namespace O365C.GraphConnector.MSLearn
{
 
    public class CreateConnection
    {
        private readonly ILearnCatalogService _learnCatalogService;
        private readonly IGraphService _graphService;
        public CreateConnection(ILearnCatalogService learnCatalogService, IGraphService graphService)
        {
            _learnCatalogService = learnCatalogService;
            _graphService = graphService;

        }
        [FunctionName("CreateConnection")]
        public async Task<IActionResult> Run(
            [HttpTrigger(AuthorizationLevel.Anonymous, "get", "post", Route = null)] HttpRequest req,
            ILogger log, ExecutionContext context)
        {
            log.LogInformation($"CreateConnection HTTP trigger function executed at: {DateTime.Now}");


            log.LogInformation("Creating connection...");
            await CreateConnectionAsync(log, context);

            return new OkObjectResult("All steps are completed");

            //string name = req.Query["name"];

            //string requestBody = await new StreamReader(req.Body).ReadToEndAsync();
            //dynamic data = JsonConvert.DeserializeObject(requestBody);
            //name = name ?? data?.name;

            //string responseMessage = string.IsNullOrEmpty(name)
            //    ? "This HTTP triggered function executed successfully. Pass a name in the query string or in the request body for a personalized response."
            //    : $"Hello, {name}. This HTTP triggered function executed successfully.";

            //return new OkObjectResult(responseMessage);
        }


        public async Task CreateConnectionAsync(ILogger log, ExecutionContext context)
        {
            _ = _graphService ?? throw new MemberAccessException("graphHttpService is null");

            try
            {
                ExternalConnection connection = await _graphService.GetConnectionAsync(ConnectionConfiguration.ConnectionID);

                if (connection == null)
                {
                    log.LogInformation("No connection was found, creating it...");
                    var filePath = Path.Combine(context.FunctionAppDirectory, "Assets", "resultLayout.json");
                    var adaptiveCard = File.ReadAllText(filePath);
                    connection = await _graphService.CreateConnectionAsync(adaptiveCard);
                    log.LogInformation("Done");
                }
                else
                {
                    log.LogInformation("Connection already exists, Skipping");
                }
            }
            catch (Exception ex)
            {
                log.LogError(ex, "Error creating connection");
                throw;
            }
        }

    }
}
