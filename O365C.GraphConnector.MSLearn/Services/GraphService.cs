//Create GrapHttpService.cs file under Services folder and add the below code   


using System;
using System.Collections.Generic;
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
using O365C.GraphConnector.MSLearn.Models;
using O365C.GraphConnector.MSLearn.Utils;

namespace O365C.GraphConnector.MSLearn.Services
{

    public interface IGraphService
    {
        Task<ExternalConnection> GetConnectionAsync(string connectionId);

        Task<ExternalConnection> CreateConnectionAsync(string adaptiveCard, string connectorId = "", string connectorTicket = "");

        Task DeleteConnectionAsync(string connectionId);

        Task<Schema> GetSchemaAsync(string connectionId);

        Task CreateSchemaAsync();
        

    }
    public class GraphService : IGraphService
    {

        private const string GraphBaseUrl = "https://graph.microsoft.com/v1.0";
        private readonly HttpClient _client;
        private readonly IAccessTokenProvider _accessTokenProvider;

        public GraphService(IHttpClientFactory httpClientFactory, IAccessTokenProvider accessTokenProvider)
        {

            _client = httpClientFactory.CreateClient();
            _accessTokenProvider = accessTokenProvider;
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
                        description = "This is a connector created for Microsoft Learn Catalog API",
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
                                        additionalProperties = adaptiveCard
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
                                        urlPattern = "/training/modules/(?<slug>[^/]+)/?",
                                    },
                                    itemId = "{slug}",
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
                                    isRetrievable = "true"
                                },
                                new
                                {
                                    name = "Roles",
                                    type = "String",
                                    isRetrievable = "true"
                                },
                                new
                                {
                                    name = "Products",
                                    type = "String",
                                    isRetrievable = "true"
                                },
                                new
                                {
                                    name = "Subjects",
                                    type = "String",
                                    isRetrievable = "true"
                                },
                                new
                                {
                                    name = "Uid",
                                    type = "String",
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
                                    isRetrievable = "true"
                                },
                                new
                                {
                                    name = "Rating",
                                    type = "String",
                                    isRetrievable = "true"
                                },
                                new
                                {
                                    name = "IconUrl",
                                    type = "String",
                                    isRetrievable = "true"
                                },
                                new
                                {
                                    name = "SocialImageUrl",
                                    type = "String",
                                    isRetrievable = "true"
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
                                    isRetrievable = "true"
                                },
                                new
                                {
                                    name = "Units",
                                    type = "String",
                                    isRetrievable = "true"
                                },
                                new
                                {
                                    name = "NumberOfUnits",
                                    type = "String",
                                    isRetrievable = "true"
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


    }
}