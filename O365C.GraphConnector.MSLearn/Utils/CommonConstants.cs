using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace O365C.GraphConnector.MSLearn.Utils
{
  public static class CommonConstants
  {
    public static string LearnCatalogApiBaseUrl { get { return "https://learn.microsoft.com"; } }
    public static string AssetsDirectoryPath { get; set; }
    public static string ResultLayoutFilePath { get; set; }
  }
}
