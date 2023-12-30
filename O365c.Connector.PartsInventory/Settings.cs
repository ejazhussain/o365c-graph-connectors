using Microsoft.Extensions.Configuration;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace O365c.Connector.PartsInventory
{
    public class Settings
    {
        public string? ClientId { get; set; }
        public string? ClientSecret { get; set; }
        public string? TenantId { get; set; }

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
