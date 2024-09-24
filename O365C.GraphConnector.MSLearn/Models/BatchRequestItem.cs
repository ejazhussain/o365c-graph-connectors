using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace O365C.GraphConnector.MSLearn.Models
{
    public class BatchRequestItem
    {
        public string Id { get; set; }
        public string Method { get; set; }
        public string Url { get; set; }
        public object Body { get; set; }
        public object Headers { get; set; }
    }
}
