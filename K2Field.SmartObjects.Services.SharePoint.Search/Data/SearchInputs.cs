using Microsoft.SharePoint.Client.Search.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace K2Field.SmartObjects.Services.SharePoint.Search.Data
{
    public class SearchInputs
    {
        public SearchInputs()
        {
            Sort = new Dictionary<string, SortDirection>();
        }

        public string Search { get; set; }
        public int? StartRow { get; set; }
        public int? RowLimit { get; set; }
        public string SortString { get; set; }
        public Dictionary<string, SortDirection> Sort { get; set; }
        public Guid SourceId { get; set; }
        public bool? EnableNicknames { get; set; }
        public bool? EnablePhonetic { get; set; }
        public string SerializedQuery { get; set; }
        public string WasGroupRestricted { get; set; }
        public string EnableInterleaving { get; set; }
        public string piPageImpression { get; set; }
        public Guid CorrelationId { get; set; }
        // processbestbits
        // showpeoplenamesuggestions
        // summary length

    }
}
