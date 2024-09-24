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
        private readonly IAccessTokenProvider _accessTokenProvider;

        public IngestContent(IGraphService graphService, ILearnCatalogService learnCatalogService, IAccessTokenProvider accessTokenProvider)
        {
            _graphService = graphService;
            _learnCatalogService = learnCatalogService;
            _accessTokenProvider = accessTokenProvider;
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

            //Get access token
            var accessToken = await _accessTokenProvider.GetAccessTokenAsync();
            if (accessToken == null)
            {
                log.LogError("Access token is null");
                return new UnauthorizedResult();
            }

            //ingest modules
            log.LogInformation("Ingesting modules...");

            for (int i = 0; i < modules.Count; i++)
            {
                var module = modules[i];
                log.LogInformation(string.Format("Loading item {0}:{1})...", i + 1,module.Title));
                 try
                {
                    await _graphService.CreateItemAsync(module, accessToken);
                    log.LogInformation("DONE");

                }
                catch (Exception ex)
                {
                    log.LogError("ERROR");
                    log.LogError(ex.Message);
                    throw;
                }
            }          


            return new OkObjectResult(modules);
        }
    }
}
