using System;
using System.IO;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using O365C.GraphConnector.MSLearn.Utils;
using O365C.GraphConnector.MSLearn.Services;
using O365C.GraphConnector.MSLearn.Models;
using System.Collections.Generic;

namespace O365C.GraphConnector.MSLearn
{
    public class IngestContent
    {
        private readonly IGraphService _graphService;
        private readonly ILearnCatalogService _learnCatalogService;
        public IngestContent(IGraphService graphService, ILearnCatalogService learnCatalogService)
        {
            _graphService = graphService;
            _learnCatalogService = learnCatalogService;
            
        }
        [FunctionName("IngestContent")]
        public async Task<IActionResult> Run(
            [HttpTrigger(AuthorizationLevel.Anonymous, "get", "post", Route = null)] HttpRequest req,
            ILogger log)
        {
            log.LogInformation($"IngestContent HTTP trigger function executed at: {DateTime.Now}");


            //load modules
            log.LogInformation("Loading modules...");
            List<Module> modules = await _learnCatalogService.GetModulesAsync();
            log.LogInformation("Loaded");


            return new OkObjectResult(modules);
        }
    }
}
