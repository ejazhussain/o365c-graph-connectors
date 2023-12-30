using System;
using System.IO;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using O365C.GraphConnector.MicrosoftLearn.Models;
using System.Linq;
using System.IdentityModel.Tokens.Jwt;
using Microsoft.IdentityModel.Tokens;

namespace O365C.GraphConnector.MicrosoftLearn
{
    public class NotificationContent
    {
        private AzureFunctionSettings _azureFunctionSettings;
        public NotificationContent(AzureFunctionSettings azureFunctionSettings)
        {
            _azureFunctionSettings = azureFunctionSettings;

        }
        [FunctionName("NotificationContent")]
        //[return: Queue("queue-content", Connection = "AzureWebJobsStorage")]
        public async Task<IActionResult> Run(
            [HttpTrigger(AuthorizationLevel.Anonymous, "get", "post", Route = null)] HttpRequest req,
            ILogger log, [Queue("queue-content", Connection = "AzureWebJobsStorage")] IAsyncCollector<string> queue)

        {
            log.LogInformation("NotificationContent was triggered.");


            log.LogInformation($"Microsoft Teams triggered our notification function...great :-)");
            var content = await new StreamReader(req.Body).ReadToEndAsync();
            log.LogInformation($"Received following payload: {content}");


            //Reterive the state value from the following payload under resourceData object            
            dynamic changeDetails = JsonConvert.DeserializeObject(content);
            string connectorState = changeDetails?.value[0]?.resourceData?.state;
            string connectorId = changeDetails?.value[0]?.resourceData?.id;



            // Grab the validationToken URL parameter
            string token = changeDetails?.validationTokens[0];

            

            // check if the token is not null or empty 
            if (string.IsNullOrEmpty(token))
            {
                log.LogError("Validation token is null or empty");
                return new BadRequestObjectResult("Validation token is null or empty");
            }
            else
            {
                log.LogInformation($"Validation token {token} received");
            }

            //Create a object of type ConnectionMessage class
            ConnectionMessage connectionMessage = new ConnectionMessage
            {
                ConnectorId = connectorId
            };

            //If the state is equal to enabled then only we need to process the notification and put the message to the queue
            if (connectorState == TargetConnectorState.Enabled)
            {
                log.LogInformation($"State value is {connectorState} and hence processing the notification");
                connectionMessage.Action = "create";
            }
            else
            {
                connectionMessage.Action = "delete";

            }

            //Validate if the connectionMessage.Action has a valid value if not then return a bad request
            if (string.IsNullOrEmpty(connectionMessage.Action))
            {
                log.LogError("Invalid action value");
                return new BadRequestObjectResult("Invalid action value");
            }

            //Put the message to the queue
            //string queueMessage = JsonConvert.SerializeObject(connectionMessage);
            //await queue.AddAsync(queueMessage);
                       

            //return object with status code 202            
            return new OkObjectResult($"{{\"status\":\"202 Accepted\"}}");

        }


    }
}
