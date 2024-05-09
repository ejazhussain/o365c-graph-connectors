using System;
using System.IO;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using O365C.GraphConnector.MSLearn.Services;
using O365C.GraphConnector.MSLearn.Utils;

namespace O365C.GraphConnector.MSLearn
{
    public class CreateSchema
    {
        private readonly IGraphService _graphService;
        public CreateSchema(IGraphService graphService)
        {
            _graphService = graphService;            
        }
        [FunctionName("CreateSchema")]
        public async Task<IActionResult> Run(
            [HttpTrigger(AuthorizationLevel.Anonymous, "get", "post", Route = null)] HttpRequest req,
            ILogger log)
        {
            log.LogInformation($"CreateSchema HTTP trigger function executed at: {DateTime.Now}");

            //Check if schema exists if not create it
            var schema = await _graphService.GetSchemaAsync(ConnectionConfiguration.ConnectionID);
            if (schema != null)
            {
                return new OkObjectResult("Schema already exists");
            }
            log.LogInformation("Creating schema...");            
            await _graphService.CreateSchemaAsync();
            log.LogInformation("Schema created");
            

            return new OkObjectResult("Schema successfully created");
        }
    }
}
