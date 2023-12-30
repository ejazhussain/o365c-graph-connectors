using Microsoft.Extensions.Configuration;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace O365.Graph.Connector.SQL
{
    public class Settings
    {
        public required string ClientId { get; set; }
        public required string ClientSecret { get; set; }
        public required string TenantId { get; set; }
        public required string ConnectionString { get; set; }

        public static Settings LoadSettings()
        {
            // Load settings
            IConfiguration config = new ConfigurationBuilder()
                .AddUserSecrets<Program>()
                .Build();

            return config.GetRequiredSection("Settings").Get<Settings>() ??
                throw new Exception("Could not load app settings. See README for configuration instructions.");
        }
    }
}
