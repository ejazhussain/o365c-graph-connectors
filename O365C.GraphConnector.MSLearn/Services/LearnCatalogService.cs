using Newtonsoft.Json;
using O365C.GraphConnector.MSLearn.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;

namespace O365C.GraphConnector.MSLearn.Services
{
    public interface ILearnCatalogService
    {
        Task<List<Module>> GetModulesAsync();
    }
    public class LearnCatalogService: ILearnCatalogService
    {
        private readonly HttpClient _client;
        private readonly string _apiBaseUrl = "https://functionapp-mslearncatalog.azurewebsites.net/api/";
        public LearnCatalogService(IHttpClientFactory httpClientFactory)
        {
            _client = httpClientFactory.CreateClient();
        }

        public async Task<List<Module>> GetModulesAsync()
        {
            List<Module> modules = new List<Module>();  
            //concatenate the base url with the path into a single string
            Uri moduleEndPoint = new Uri(_apiBaseUrl + "modules");

            var response = await _client.GetAsync(moduleEndPoint);
            if(response.EnsureSuccessStatusCode().IsSuccessStatusCode)
            {
                var jsonText = await response.Content.ReadAsStringAsync();
                modules = JsonConvert.DeserializeObject<List<Module>>(jsonText);
            }

            return modules;
        }
    }
}
