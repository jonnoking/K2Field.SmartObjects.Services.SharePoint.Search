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
            Sort = new Dictionary<string, string>();
            FileExtensions = new List<string>();
        }

        public string Search { get; set; }

        public string Site { get; set; } // to limit scope
        public string SiteUrl { get; set; }
        public string Properties { get; set; }
        public int? StartRow { get; set; }
        public int? RowLimit { get; set; }
        public string SortString { get; set; }
        public Dictionary<string, string> Sort { get; set; }
        public Guid SourceId { get; set; }
        public bool? EnableStemming { get; set; }
        public bool? TrimDuplicates { get; set; }
        public bool? ProcessPersonal { get; set; }
        public bool? ProcessBestBets { get; set; }
        public bool? EnableQueryRules { get; set; }


        public bool? EnableNicknames { get; set; }
        public bool? EnablePhonetic { get; set; }

        public List<string> FileExtensions { get; set; }
        public string FileExtensionsString { get; set; }


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
