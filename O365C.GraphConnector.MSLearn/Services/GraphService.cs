//Create GrapHttpService.cs file under Services folder and add the below code   


using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Graph.Models.ExternalConnectors;
using Microsoft.Identity.Client;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using O365C.GraphConnector.MSLearn.Helpers;
using O365C.GraphConnector.MSLearn.Models;
using O365C.GraphConnector.MSLearn.Utils;

namespace O365C.GraphConnector.MSLearn.Services
{

    public interface IGraphService
    {
        Task<ExternalConnection> GetConnectionAsync(string connectionId);

        Task<ExternalConnection> CreateConnectionAsync(string adaptiveCard, string connectorId = "", string connectorTicket = "");

        Task UpdateConnectionAsync(string adaptiveCard, string connectorId);

        Task DeleteConnectionAsync(string connectionId);

        Task<Schema> GetSchemaAsync(string connectionId);

        Task CreateSchemaAsync();

        Task CreateItemAsync(Module module, string accessToken);



    }
    public class GraphService : IGraphService
    {

        private const string GraphBaseUrl = "https://graph.microsoft.com/v1.0";
        private readonly HttpClient _client;
        private readonly IAccessTokenProvider _accessTokenProvider;
        private readonly AzureFunctionSettings _azureFunctionSettings;


        public GraphService(IHttpClientFactory httpClientFactory, IAccessTokenProvider accessTokenProvider, AzureFunctionSettings azureFunctionSettings)
        {

            _client = httpClientFactory.CreateClient();
            _accessTokenProvider = accessTokenProvider;
            _azureFunctionSettings = azureFunctionSettings;
        }

        public async Task<ExternalConnection> GetConnectionAsync(string connectionId)
        {

            try
            {
                //Get Access Token
                var accessToken = await _accessTokenProvider.GetAccessTokenAsync();
                if (accessToken != null)
                {
                    //Get Connection
                    string endpoint = $"{GraphBaseUrl}/external/connections/{connectionId}";
                    using (var request = new HttpRequestMessage(HttpMethod.Get, endpoint))
                    {
                        //Headers
                        request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
                        request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));


                        var response = _client.SendAsync(request).Result;

                        if (response.IsSuccessStatusCode)
                        {
                            var stringResult = response.Content.ReadAsStringAsync().Result;
                            var result = JsonConvert.DeserializeObject<ExternalConnection>(stringResult);
                            return result;
                        }
                        else
                        {
                            return null;
                        }

                    }
                }
                else
                {
                    throw new Exception("Error getting access token");
                }

            }
            catch (Exception ex)
            {
                throw new Exception($"Error getting connection: {ex.Message}");
            }
        }

        public async Task<ExternalConnection> CreateConnectionAsync(string adaptiveCard, string connectorId = "", string connectorTicket = "")
        {
            try
            {
                var accessToken = await _accessTokenProvider.GetAccessTokenAsync();
                string endpoint = $"{GraphBaseUrl}/external/connections";
                var layout = JsonConvert.DeserializeObject<Dictionary<string, object>>(adaptiveCard);


                using (var request = new HttpRequestMessage(HttpMethod.Post, endpoint))
                {
                    //Headers
                    request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
                    request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                    if (!string.IsNullOrEmpty(connectorTicket))
                    {
                        request.Headers.Add("GraphConnectors-Ticket", connectorTicket);
                    }


                    var payload = new
                    {
                        description = "The Microsoft Learn Catalog API provides details on Microsoft Learn's training modules (e.g., title, summary, url, path, products, roles, etc covered). Organizations use this API to access and reference this content throughout their workflows. Users can search for modules by title.",
                        id = ConnectionConfiguration.ConnectionID,
                        name = "Microsoft Learn Connector",
                        connectorId = !string.IsNullOrEmpty(connectorId) ? connectorId : null,
                        searchSettings = new
                        {
                            searchResultTemplates = new[]
                            {
                                new
                                {
                                    id = "MsLearnAPISrc",
                                    priority = 1,
                                     layout = new
                                    {
                                        Schema = "https://adaptivecards.io/schemas/adaptive-card.json",
                                        type = "AdaptiveCard",
                                        version = "1.3",
                                        body = layout["body"]
                                    }
                                }
                            }
                        },
                        activitySettings = new
                        {

                            urlToItemResolvers = new[]
                            {
                                new
                                {
                                    @odata_type = "#microsoft.graph.externalConnectors.itemIdResolver",
                                    urlMatchInfo = new
                                    {
                                        baseUrls = new[] { CommonConstants.LearnCatalogApiBaseUrl },
                                        urlPattern = "/learn/modules/(?<moduleId>[^/]+)"
                                    },
                                    itemId = "{moduleId}",
                                    priority = 1
                                }
                            }
                        }
                    };

                    string jsonContent = JsonConvert.SerializeObject(payload);

                    jsonContent = jsonContent.ToString().Replace("odata_type", "@odata.type");

                    request.Content = new StringContent(jsonContent, Encoding.UTF8, "application/json");

                    var response = _client.SendAsync(request).Result;

                    if (response.IsSuccessStatusCode)
                    {
                        var stringResult = response.Content.ReadAsStringAsync().Result;
                        var result = JsonConvert.DeserializeObject<ExternalConnection>(stringResult);
                        return result;
                    }
                    else
                    {
                        throw new Exception($"Error creating connection: {response.ReasonPhrase}");
                    }

                }
            }
            catch (Exception ex)
            {
                //Log exception
                throw new Exception($"Error creating connection: {ex.Message}");

            }
        }

        public async Task UpdateConnectionAsync(string adaptiveCard, string connectorId)
        {
            try
            {
                var accessToken = await _accessTokenProvider.GetAccessTokenAsync();
                string endpoint = $"{GraphBaseUrl}/external/connections/{connectorId}";
                var layout = JsonConvert.DeserializeObject<Dictionary<string, object>>(adaptiveCard);


                using (var request = new HttpRequestMessage(HttpMethod.Patch, endpoint))
                {
                    //Headers
                    request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
                    request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));


                    var payload = new
                    {
                        description = "The Microsoft Learn Catalog API provides details on Microsoft Learn's training modules (e.g., title, summary, url, path, products, roles, etc covered). Organizations use this API to access and reference this content throughout their workflows. Users can search for modules by title.",
                        searchSettings = new
                        {
                            searchResultTemplates = new[]
                            {
                                new
                                {
                                    id = "MsLearnAPISrc",
                                    priority = 1,
                                    layout = new
                                    {
                                        Schema = "https://adaptivecards.io/schemas/adaptive-card.json",
                                        type = "AdaptiveCard",
                                        version = "1.3",
                                        body = layout["body"]
                                    }
                                    
                                }
                            }
                        },
                        activitySettings = new
                        {

                            urlToItemResolvers = new[]
                            {
                                new
                                {
                                    @odata_type = "#microsoft.graph.externalConnectors.itemIdResolver",
                                    urlMatchInfo = new
                                    {
                                        baseUrls = new[] { CommonConstants.LearnCatalogApiBaseUrl },
                                        urlPattern = "/[^/]+/(?<moduleId>[^/]+)$"
                                    },
                                    itemId = "{moduleId}",
                                    priority = 1
                                }
                            }
                        }

                    };

                    string jsonContent = JsonConvert.SerializeObject(payload);
                    jsonContent = jsonContent.ToString().Replace("odata_type", "@odata.type");

                    request.Content = new StringContent(jsonContent, Encoding.UTF8, "application/json");

                    var response = _client.SendAsync(request).Result;

                }
            }
            catch (Exception ex)
            {
                //Log exception
                throw new Exception($"Error creating connection: {ex.Message}");

            }
        }
        public async Task DeleteConnectionAsync(string connectionId)
        {
            try
            {

                var accessToken = await _accessTokenProvider.GetAccessTokenAsync();

                string endpoint = $"{GraphBaseUrl}/external/connections/{connectionId}";

                using (var request = new HttpRequestMessage(HttpMethod.Delete, endpoint))
                {
                    request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);

                    var response = _client.SendAsync(request).Result;
                }
            }
            catch (Exception ex)
            {
                //Log exception
                throw new Exception($"Error deleting connection: {ex.Message}");

            }

        }
        public async Task<Schema> GetSchemaAsync(string connectionId)
        {
            try
            {
                var accessToken = await _accessTokenProvider.GetAccessTokenAsync();

                string endpoint = $"{GraphBaseUrl}/external/connections/{connectionId}/schema";

                using (var request = new HttpRequestMessage(HttpMethod.Get, endpoint))
                {
                    request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);

                    var response = _client.SendAsync(request).Result;
                    if (response.IsSuccessStatusCode)
                    {
                        var stringResult = response.Content.ReadAsStringAsync().Result;
                        var result = JsonConvert.DeserializeObject<Schema>(stringResult);
                        return result;
                    }
                    else
                    {
                        //If not found then return null
                        if (response.StatusCode == System.Net.HttpStatusCode.NotFound)
                        {
                            return null;
                        }
                        else
                        {
                            throw new Exception($"Error getting schema: {response.ReasonPhrase}");
                        }
                    }
                }

            }
            catch (Exception ex)
            {
                //Log exception
                throw new Exception($"Error getting schema: {ex.Message}");

            }
        }
        public async Task CreateSchemaAsync()
        {

            try
            {
                var accessToken = await _accessTokenProvider.GetAccessTokenAsync();
                string endpoint = $"{GraphBaseUrl}/external/connections/{ConnectionConfiguration.ConnectionID}/schema";
                using var request = new HttpRequestMessage(HttpMethod.Patch, endpoint);
                //Headers
                request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
                request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));


                var payload = new
                {
                    baseType = "microsoft.graph.externalItem",
                    properties = new object[]
                       {
                                new
                                {
                                    name = "Summary",
                                    type = "String",
                                    isQueryable = "true",
                                    isSearchable = "true",
                                    isRetrievable = "true"

                                },
                                new
                                {
                                    name = "Levels",
                                    type = "String",
                                    isSearchable = "true",
                                    isRetrievable = "true"
                                },
                                new
                                {
                                    name = "Roles",
                                    type = "String",
                                    isSearchable = "true",
                                    isRetrievable = "true"
                                },
                                new
                                {
                                    name = "Products",
                                    type = "String",
                                    isSearchable = "true",
                                    isRetrievable = "true"
                                },
                                new
                                {
                                    name = "Subjects",
                                    type = "String",
                                    isSearchable = "true",
                                    isRetrievable = "true"
                                },
                                new
                                {
                                    name = "Uid",
                                    type = "String",
                                    isSearchable = "true",
                                    isRetrievable = "true"
                                },
                                new
                                {
                                    name = "Title",
                                    type = "String",
                                    isQueryable = "true",
                                    isSearchable = "true",
                                    isRetrievable = "true",
                                    labels = new[] { "Title" }
                                },
                                new
                                {
                                    name = "Duration",
                                    type = "String",
                                    isSearchable = "true",
                                    isRetrievable = "true"


                                },
                                new
                                {
                                    name = "Rating",
                                    type = "String",
                                    isSearchable = "true",
                                    isRetrievable = "true"


                                },
                                new
                                {
                                    name = "IconUrl",
                                    type = "String",
                                    isRetrievable = "true",
                                    isSearchable = "true",
                                    labels = new[] { "IconUrl" }

                                },
                                new
                                {
                                    name = "SocialImageUrl",
                                    type = "String",
                                    isRetrievable = "true",
                                    isSearchable = "true",

                                },
                                 new
                                {
                                    name = "LastModified",
                                    type = "DateTime",
                                    isQueryable = "true",
                                    isRefinable = "true",
                                    isRetrievable = "true",
                                    labels = new[] { "LastModifiedDateTime" }
                                },
                                new
                                {
                                    name = "Path",
                                    type = "String",
                                    isQueryable = "true",
                                    isRetrievable = "true",
                                    isSearchable = "true",
                                    Labels = new[] { "url" }

                                },
                                new
                                {
                                    name = "Units",
                                    type = "String",
                                    isRetrievable = "true",
                                    isSearchable = "true",
                                },
                                new
                                {
                                    name = "NumberOfUnits",
                                    type = "String",
                                    isRetrievable = "true",
                                    isSearchable = "true",
                                }

                        }
                };


                string jsonContent = JsonConvert.SerializeObject(payload);

                request.Content = new StringContent(jsonContent, Encoding.UTF8, "application/json");

                var response = _client.SendAsync(request).Result;

                if (response.IsSuccessStatusCode)
                {
                    // The operation ID is contained in the Location header returned
                    // in the response
                    var operationId = response.Headers.Location?.Segments.Last() ??
                        throw new Exception("Could not get operation ID from Location header");
                    await WaitForOperationToCompleteAsync(accessToken, ConnectionConfiguration.ConnectionID, operationId);

                }
                else
                {
                    //throw new Exception($"Error creating schema: {response.ReasonPhrase}");
                    throw new ServiceException("creating schema failed", response.Headers, (int)response.StatusCode);
                }
            }
            catch (Exception ex)
            {
                //Log exception
                throw new Exception($"Error creating schema: {ex.Message}");

            }
        }

        public async Task CreateItemAsync(Module module, string accessToken)
        {

            try
            {
                var itemId = Regex.Replace(module.Uid, "[^a-zA-Z0-9-]", "");
                string endpoint = $"{GraphBaseUrl}/external/connections/{ConnectionConfiguration.ConnectionID}/items/{itemId}";
                using var request = new HttpRequestMessage(HttpMethod.Put, endpoint);
                //Headers
                request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
                request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

                //payload
                var payload = new
                {
                    id = itemId,
                    acl = new object[]
                {
                        new
                        {
                            type = "everyone",
                            value = "everyone",
                            accessType = "grant"
                        }
                },
                    properties = new
                    {
                        Summary = module.Summary ?? string.Empty,
                        Levels = string.Join(",", module.Levels ?? new List<string>()),
                        Roles = string.Join(',', module.Roles ?? new List<string>()),
                        Products = string.Join(",", module.Products ?? new List<string>()),
                        Subjects = string.Join(",", module.Subjects ?? new List<string>()),
                        Uid = module.Uid ?? "",
                        Title = module.Title ?? "",
                        Duration = module.Duration.ToString() ?? "",
                        Rating = module.Rating != null ? ModuleHelper.GetStarRating((int)module.Rating.Average) : "",
                        IconUrl = module.IconUrl ?? "",
                        SocialImageUrl = module.SocialImageUrl ?? "",
                        LastModified = module.LastModified,
                        Path = module.Path ?? "",
                        Units = string.Join(",", module.Units ?? new List<string>()),
                        NumberOfUnits = module.NumberOfUnits.ToString() ?? ""
                    },
                    content = new
                    {
                        value = module.Summary ?? "",
                        type = "text"
                    },
                    activities = new[]
                    {
                        new
                        {
                            OdataType = "#microsoft.graph.externalConnectors.externalActivity",
                            Type = ExternalActivityType.Modified,
                            StartDateTime = DateTimeOffset.Parse(module.LastModified.ToString()),
                            performedBy = new
                            {
                                type = "user",
                                id = _azureFunctionSettings.UserId,
                            }

                        }
                    }
                };

                string jsonContent = JsonConvert.SerializeObject(payload);
                request.Content = new StringContent(jsonContent, Encoding.UTF8, "application/json");
                var response = await _client.SendAsync(request);

            }
            catch (Exception ex)
            {
                //Log exception
                throw new Exception($"Error creating item: {ex.Message}");
            }


        }

        private async Task WaitForOperationToCompleteAsync(string accessToken, string connectionId, string operationId)
        {
            do
            {

                string endpoint = $"{GraphBaseUrl}/external/connections/{connectionId}/operations/{operationId}";

                using var request = new HttpRequestMessage(HttpMethod.Get, endpoint);
                //Headers
                request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);

                var response = _client.SendAsync(request).Result;

                if (response.IsSuccessStatusCode)
                {
                    var stringResult = response.Content.ReadAsStringAsync().Result;
                    var result = JsonConvert.DeserializeObject<ConnectionOperation>(stringResult);

                    if (result.Status == ConnectionOperationStatus.Completed)
                    {
                        return;
                    }
                    else if (result.Status == ConnectionOperationStatus.Failed)
                    {
                        throw new ServiceException($"Schema operation failed: {result.Error?.Code} {result.Error?.Message}");
                    }
                }
                else
                {
                    throw new Exception($"Error getting operation status: {response.ReasonPhrase}");
                }
                // Wait 5 seconds and check again
                await Task.Delay(5000);
            } while (true);

        }

        private void AddItemActivities(string accessToken, string connectionId, string itemId, Module module)
        {
            try
            {
                string endpoint = $"{GraphBaseUrl}/external/connections/{connectionId}/items/{itemId}/activities";
                using var request = new HttpRequestMessage(HttpMethod.Post, endpoint);
                //Headers
                request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
                request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

                //payload
                var payload = new
                {
                    activities = new object[]
        {
            new
            {
                OdataType = "#microsoft.graph.externalConnectors.externalActivity",
                 Type =  ExternalActivityType.Modified,
                StartDateTime = DateTimeOffset.Parse(module.LastModified.ToString()),
                performedBy = new
                {
                    type = "user",
                    id = _azureFunctionSettings.UserId,
                }
            }
        }

                };

                string jsonContent = JsonConvert.SerializeObject(payload);
                request.Content = new StringContent(jsonContent, Encoding.UTF8, "application/json");
                var response = _client.SendAsync(request).Result;

                if (!response.IsSuccessStatusCode)
                {
                    throw new Exception($"Error adding item activities: {response.ReasonPhrase}");
                }
            }
            catch (Exception ex)
            {
                //Log exception
                throw new Exception($"Error adding item activities: {ex.Message}");
            }
        }

        private async Task<string> SendBatchRequest(string accessToken, string batchContent)
        {
            using (var httpClient = new HttpClient())
            {
                httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
                httpClient.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

                var content = new StringContent(batchContent, Encoding.UTF8, "application/json");
                var response = await httpClient.PostAsync("https://graph.microsoft.com/v1.0/$batch", content);

                if (response.IsSuccessStatusCode)
                {
                    return await response.Content.ReadAsStringAsync();
                }
                else
                {
                    throw new Exception($"Error sending batch request: {response.StatusCode}");
                }
            }
        }
    }


}