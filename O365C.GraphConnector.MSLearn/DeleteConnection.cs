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

    public class DeleteConnection
    {
        
        private readonly IGraphService _graphService;
        public DeleteConnection(IGraphService graphService)
        {            
            _graphService = graphService;

        }
        [FunctionName("DeleteConnection")]
        public async Task<IActionResult> Run(
            [HttpTrigger(AuthorizationLevel.Anonymous, "get", "post", Route = null)] HttpRequest req,
            ILogger log, ExecutionContext context)
        {
            log.LogInformation($"DeleteConnection HTTP trigger function executed at: {DateTime.Now}");


            log.LogInformation("Deleting connection...");
            await _graphService.DeleteConnectionAsync(ConnectionConfiguration.ConnectionID);
            log.LogInformation("Deleted");
                        

            return new OkObjectResult("Connection successfully deleted");

        }        

    }
}
