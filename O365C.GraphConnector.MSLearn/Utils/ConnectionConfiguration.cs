using Microsoft.Graph.Models.ExternalConnectors;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Text.Json;
using System.IO;
using Microsoft.Graph.Models;

namespace O365C.GraphConnector.MSLearn.Utils
{
    internal class ConnectionConfiguration
    {
        public static string ConnectionID = "MsLearnConnector";
                
        public static Schema Schema
        {
            get
            {
                return new Schema()
                {
                    BaseType = "microsoft.graph.externalItem",
                    Properties = new List<Property>() {
                        new Property {
                            Name = "Summary",
                            Type = PropertyType.String,
                            IsQueryable = true,
                            IsSearchable = true,
                            IsRetrievable = true,
                        },
                        new Property {
                            Name = "Levels",
                            Type = PropertyType.String,
                            IsRetrievable = true
                        },
                        new Property {
                            Name = "Roles",
                            Type = PropertyType.String,
                            IsRetrievable = true
                        },
                        new Property {
                            Name = "Products",
                            Type = PropertyType.String,
                            IsRetrievable = true
                        },
                        new Property {
                            Name = "Subjects",
                            Type = PropertyType.String,
                            IsRetrievable = true
                        },
                        new Property {
                            Name = "Uid",
                            Type = PropertyType.String,
                            IsRetrievable = true
                        },
                        new Property {
                            Name = "Title",
                            Type = PropertyType.String,
                            IsQueryable = true,
                            IsSearchable = true,
                            IsRetrievable = true,
                             Labels = new() { Label.Title }
                        },
                         new Property {
                            Name = "Duration",
                            Type = PropertyType.String,
                            IsRetrievable = true
                        },
                          new Property {
                            Name = "Rating",
                            Type = PropertyType.String,
                            IsRetrievable = true
                        },
                        new Property {
                            Name = "IconUrl",
                            Type = PropertyType.String,
                            IsRetrievable = true
                        },
                        new Property {
                            Name = "SocialImageUrl",
                            Type = PropertyType.String,
                            IsRetrievable = true
                        },
                        new Property {
                            Name = "LastModified",
                            Type = PropertyType.DateTime,
                            IsQueryable = true,
                            IsRefinable = true,
                            IsRetrievable = true,
                            Labels = new() { Label.LastModifiedDateTime }
                        },
                         new Property {
                            Name = "Path",
                            Type = PropertyType.String,
                            IsRetrievable = true,
                            Labels = new() { Label.Url }
                        },
                        new Property {
                            Name = "Units",
                            Type = PropertyType.String,
                            IsRetrievable = true
                        },
                        new Property {
                            Name = "NumberOfUnits",
                            Type = PropertyType.String,
                            IsRetrievable = true,
                        },
                    }
                };
            }
        }
    }
}

