using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace K2Field.SmartObjects.Services.SharePoint.Search.Data
{
    public class SearchResultsSerialized
    {
        public SearchInputs Inputs { get; set; }
        public int? ResultRows { get; set; }
        public int? TotalRows { get; set; }
        public int? ExecutionTime { get; set; }
        public string ResultTitle { get; set; }
        public string ResultTitleUrl { get; set; }
        public string SpellingSuggestions { get; set; } // ???
        public string TableType { get; set; }
        public IEnumerable<IDictionary<string, object>> SearchResults { get; set; }

    }
}
