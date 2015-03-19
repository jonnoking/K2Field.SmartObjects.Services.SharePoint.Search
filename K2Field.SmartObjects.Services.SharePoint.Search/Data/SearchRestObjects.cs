using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace K2Field.SmartObjects.Services.SharePoint.Search.Data
{
    class SearchRestObjects
    {
    }


    
public class RESTSearchResults
{
public string odatametadata { get; set; }
public int ElapsedTime { get; set; }
public PrimaryQueryResult PrimaryQueryResult { get; set; }
public SearchProperty[] Properties { get; set; }
public object[] SecondaryQueryResults { get; set; }
public string SpellingSuggestion { get; set; }
public object[] TriggeredRules { get; set; }
}

public class PrimaryQueryResult
{
public object[] CustomResults { get; set; }
public string QueryId { get; set; }
public string QueryRuleId { get; set; }
public object RefinementResults { get; set; }
public RelevantResults RelevantResults { get; set; }
public object SpecialTermResults { get; set; }
}

public class RelevantResults
{
public object GroupTemplateId { get; set; }
public object ItemTemplateId { get; set; }
public SearchProperty[] Properties { get; set; }
public string ResultTitle { get; set; }
public string ResultTitleUrl { get; set; }
public int RowCount { get; set; }
public ResultTable Table { get; set; }
public int TotalRows { get; set; }
public int TotalRowsIncludingDuplicates { get; set; }
}

public class ResultTable
{
public ResultRow[] Rows { get; set; }
}

public class ResultRow
{
public ResultCell[] Cells { get; set; }
}

public class ResultCell
{
public string Key { get; set; }
public object Value { get; set; }
public string ValueType { get; set; }
}

public class SearchProperty
{
public string Key { get; set; }
public string Value { get; set; }
public string ValueType { get; set; }
}


}
