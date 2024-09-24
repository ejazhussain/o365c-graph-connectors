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
 
    public class ManageConnection
    {
        private readonly ILearnCatalogService _learnCatalogService;
        private readonly IGraphService _graphService;
        public ManageConnection(ILearnCatalogService learnCatalogService, IGraphService graphService)
        {
            _learnCatalogService = learnCatalogService;
            _graphService = graphService;

        }
        [FunctionName("ManageConnection")]
        public async Task<IActionResult> Run(
            [HttpTrigger(AuthorizationLevel.Anonymous, "get", "post", Route = null)] HttpRequest req,
            ILogger log, ExecutionContext context)
        {
            log.LogInformation($"ManageConnection HTTP trigger function executed at: {DateTime.Now}");

            string action = req.Query["action"];
            string requestBody = await new StreamReader(req.Body).ReadToEndAsync();
            dynamic data = JsonConvert.DeserializeObject(requestBody);
            action = action ?? data?.action;

            if (action.ToLower() == "create")
            {
                log.LogInformation("Creating connection...");
                await CreateConnectionAsync(log, context);
            }
            else if (action.ToLower() == "update")
            {
                log.LogInformation("Updating connection...");
                await UpdateConnectionAsync(log, context);
            }
            else
            {
                log.LogInformation("No action specified, skipping...");
            }
            

            return new OkObjectResult("All steps are completed");
           

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

                var filePath = Path.Combine(context.FunctionAppDirectory, "Assets", "resultLayout.json");
                var adaptiveCard = File.ReadAllText(filePath);                

                if (connection == null)
                {
                    log.LogInformation("No connection was found, creating it...");                   
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

        public async Task UpdateConnectionAsync(ILogger log, ExecutionContext context)
        {
            _ = _graphService ?? throw new MemberAccessException("graphHttpService is null");

            try
            {
                ExternalConnection connection = await _graphService.GetConnectionAsync(ConnectionConfiguration.ConnectionID);

                var filePath = Path.Combine(context.FunctionAppDirectory, "Assets", "resultLayout.json");
                var adaptiveCard = File.ReadAllText(filePath);

                if (connection == null)
                {
                    log.LogInformation("No connection was found, Skipping...");                   
                }
                else
                {
                    log.LogInformation("Connection already exists, Updating");
                    await _graphService.UpdateConnectionAsync(adaptiveCard, connection.Id);
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
