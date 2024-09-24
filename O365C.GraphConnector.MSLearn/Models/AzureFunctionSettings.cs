using System.Security.Cryptography.X509Certificates;

namespace O365C.GraphConnector.MSLearn.Models
{ 
    public class AzureFunctionSettings
    {        
        public string TenantId { get; set; }        
        public string ClientId { get; set; }            
        public string ClientSecret { get; set; }        
        public string UserId { get; set; }
    }
}
