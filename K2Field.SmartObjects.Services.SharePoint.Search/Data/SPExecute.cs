using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SourceCode.SmartObjects.Services.ServiceSDK;
using SourceCode.SmartObjects.Services.ServiceSDK.Objects;
using SourceCode.SmartObjects.Services.ServiceSDK.Types;
using Newtonsoft.Json;
using System.Net;
using System.IO;
using System.Globalization;

namespace K2Field.SmartObjects.Services.SharePoint.Search.Data
{
    public class SPExecute
    {
        private ServiceAssemblyBase serviceBroker = null;
        private Configuration Configuration { get; set; }

        public SPExecute(ServiceAssemblyBase serviceBroker, Configuration configuration)
        {
            // Set local serviceBroker variable.
            this.serviceBroker = serviceBroker;
            this.Configuration = configuration;
        }

        public void ExecuteSearch(Property[] inputs, RequiredProperties required, Property[] returns, MethodType methodType, ServiceObject serviceObject)
        {
            serviceObject.Properties.InitResultTable();
            System.Data.DataRow dr;
            try
            {
                RESTSearchResultsSerialized SerializedResults = null;

                // if deserializesearchresults
                if (serviceObject.Methods[0].Name.StartsWith("deserialize"))
                {
                    Property SerializedProp = inputs.Where(p => p.Name.Equals("serializedresults", StringComparison.OrdinalIgnoreCase)).First();
                    string json = string.Empty;
                    json = SerializedProp.Value.ToString();

                    SerializedResults = JsonConvert.DeserializeObject<RESTSearchResultsSerialized>(json.Trim());

                    if (string.IsNullOrWhiteSpace(json) || SerializedResults == null)
                    {
                        throw new Exception("Failed to deserialize search results");
                    }
                }

                if (serviceObject.Methods[0].Name.Equals("spsearch") || serviceObject.Methods[0].Name.Equals("spsearchdocuments"))
                {
                    // if Search
                    SerializedResults = Utilities.BrokerUtils.ExecuteSharePointSearch(inputs, required, Configuration, serviceBroker);
                }

                if (serviceObject.Methods[0].Name.Equals("spsearchraw"))
                {
                    // if Search Raw Read
                    SerializedResults = Utilities.BrokerUtils.ExecuteSharePointSearchRaw(inputs, required, Configuration, serviceBroker);
                }


                if (serviceObject.Methods[0].Name.Equals("spsearchusers"))
                {
                    // if Search
                    SerializedResults = Utilities.BrokerUtils.ExecuteSharePointUserSearch(inputs, required, Configuration, serviceBroker);
                }


                if (SerializedResults != null)
                {
                    if (SerializedResults.SearchResults == null || SerializedResults.SearchResults.Rows == null)
                    {
                        SerializedResults.SearchResults = new ResultTable();

                        //SerializedResults.SearchResults.Rows.Add(new ResultRow());
                    }
                    
                    foreach (ResultRow result in SerializedResults.SearchResults.Rows)
                    {
                        dr = serviceBroker.ServicePackage.ResultTable.NewRow();

                        if (!string.IsNullOrWhiteSpace(SerializedResults.Inputs.Search))
                        {
                            dr["search"] = SerializedResults.Inputs.Search;
                        }

                        if (!string.IsNullOrWhiteSpace(SerializedResults.Inputs.SiteUrl))
                        {
                            dr["searchsiteurl"] = SerializedResults.Inputs.Search;
                        }

                        if (!string.IsNullOrWhiteSpace(SerializedResults.Inputs.FileExtensionsString))
                        {
                            dr["fileextensionsfilter"] = SerializedResults.Inputs.FileExtensionsString;
                        }

                        if (SerializedResults.Inputs.SourceId != null && SerializedResults.Inputs.SourceId != Guid.Empty)
                        {
                            dr["sourceid"] = SerializedResults.Inputs.SourceId;
                        }

                        if (!string.IsNullOrWhiteSpace(SerializedResults.Inputs.SortString))
                        {
                            dr["sort"] = SerializedResults.Inputs.SortString;
                        }

                        if (SerializedResults.Inputs.StartRow.HasValue && SerializedResults.Inputs.StartRow.Value > -1)
                        {
                            dr["startrow"] = SerializedResults.Inputs.StartRow.Value;
                        }

                        if (SerializedResults.Inputs.RowLimit.HasValue && SerializedResults.Inputs.RowLimit.Value > 0)
                        {
                            dr["rowlimit"] = SerializedResults.Inputs.RowLimit.Value;
                        }

                        if (!string.IsNullOrWhiteSpace(SerializedResults.Inputs.Properties))
                        {
                            dr["properties"] = SerializedResults.Inputs.Properties;
                        }

                        if (SerializedResults.Inputs.EnableStemming.HasValue && SerializedResults.Inputs.EnableStemming.Value)
                        {
                            dr["enablestemming"] = SerializedResults.Inputs.EnableStemming.Value;
                        }

                        if (SerializedResults.Inputs.TrimDuplicates.HasValue && SerializedResults.Inputs.TrimDuplicates.Value)
                        {
                            dr["trimduplicates"] = SerializedResults.Inputs.TrimDuplicates.Value;
                        }

                        if (SerializedResults.Inputs.EnableQueryRules.HasValue && SerializedResults.Inputs.EnableQueryRules.Value)
                        {
                            dr["enablequeryrules"] = SerializedResults.Inputs.EnableQueryRules.Value;
                        }

                        if (SerializedResults.Inputs.ProcessBestBets.HasValue && SerializedResults.Inputs.ProcessBestBets.Value)
                        {
                            dr["processbestbets"] = SerializedResults.Inputs.ProcessBestBets.Value;
                        }

                        if (SerializedResults.Inputs.ProcessPersonal.HasValue && SerializedResults.Inputs.ProcessPersonal.Value)
                        {
                            dr["processpersonal"] = SerializedResults.Inputs.ProcessPersonal.Value;
                        }

                        if (SerializedResults.Inputs.EnableNicknames.HasValue && SerializedResults.Inputs.EnableNicknames.Value)
                        {
                            dr["enablenicknames"] = SerializedResults.Inputs.EnableNicknames.Value;
                        }

                        if (SerializedResults.Inputs.EnablePhonetic.HasValue && SerializedResults.Inputs.EnablePhonetic.Value)
                        {
                            dr["enablephonetic"] = SerializedResults.Inputs.EnablePhonetic.Value;
                        }

                        if (SerializedResults.ExecutionTime.HasValue)
                        {
                            dr["executiontime"] = SerializedResults.ExecutionTime.Value;
                        }

                        if (SerializedResults.ResultRows.HasValue)
                        {
                            dr["resultrows"] = SerializedResults.ResultRows.Value;
                        }
                        if (SerializedResults.TotalRows.HasValue)
                        {
                            dr["totalrows"] = SerializedResults.TotalRows.Value;
                        }

                        List<string> missingprops = new List<string>();
                        foreach (ResultCell cell in result.Cells)
                        {
                            if (dr.Table.Columns.Contains(cell.Key.ToLower()))
                            {
                                if (cell.Value != null)
                                {
                                    dr[cell.Key.ToLower()] = cell.Value;
                                }
                            }
                            else
                            {
                                missingprops.Add(cell.Key);
                            }
                        }

                        dr["responsestatus"] = ResponseStatus.Success;
                        serviceBroker.ServicePackage.ResultTable.Rows.Add(dr);
                    }
                }
                else
                {
                    throw new Exception("No results returned.");
                }

            }
            catch (Exception ex)
            {
                dr = serviceBroker.ServicePackage.ResultTable.NewRow();
                dr["responsestatus"] = ResponseStatus.Error;
                dr["responsestatusdescription"] = ex.Message;
                serviceBroker.ServicePackage.ResultTable.Rows.Add(dr);
            }

            //serviceObject.Properties.BindPropertiesToResultTable();
        }

        public void ExecuteSearchRead(Property[] inputs, RequiredProperties required, Property[] returns, MethodType methodType, ServiceObject serviceObject)
        {
            serviceObject.Properties.InitResultTable();

            try
            {
                RESTSearchResultsSerialized SerializedResults = new RESTSearchResultsSerialized();


                if (serviceObject.Methods[0].Name.Equals("spsearchread") || serviceObject.Methods[0].Name.Equals("spsearchdocumentsread"))
                {
                    SerializedResults = Utilities.BrokerUtils.ExecuteSharePointSearch(inputs, required, Configuration, serviceBroker);
                }

                if (serviceObject.Methods[0].Name.Equals("spsearchrawread"))
                {
                    SerializedResults = Utilities.BrokerUtils.ExecuteSharePointSearchRaw(inputs, required, Configuration, serviceBroker);
                }

                if (serviceObject.Methods[0].Name.Equals("spsearchusersread"))
                {
                    SerializedResults = Utilities.BrokerUtils.ExecuteSharePointUserSearch(inputs, required, Configuration, serviceBroker);
                }


                if (SerializedResults != null)
                {
                    if (!string.IsNullOrWhiteSpace(SerializedResults.Inputs.Search))
                    {
                        returns.Where(p => p.Name.Equals("search", StringComparison.OrdinalIgnoreCase)).First().Value = SerializedResults.Inputs.Search;
                    }

                    if (!string.IsNullOrWhiteSpace(SerializedResults.Inputs.SiteUrl))
                    {
                        returns.Where(p => p.Name.Equals("searchsiteurl", StringComparison.OrdinalIgnoreCase)).First().Value = SerializedResults.Inputs.SiteUrl;
                    }

                    if (!string.IsNullOrWhiteSpace(SerializedResults.Inputs.FileExtensionsString))
                    {
                        returns.Where(p => p.Name.Equals("fileextensionsfilter", StringComparison.OrdinalIgnoreCase)).First().Value = SerializedResults.Inputs.FileExtensionsString;
                    }

                    if (SerializedResults.Inputs.SourceId != null && SerializedResults.Inputs.SourceId != Guid.Empty)
                    {
                        returns.Where(p => p.Name.Equals("sourceid", StringComparison.OrdinalIgnoreCase)).First().Value = SerializedResults.Inputs.SourceId;
                    }

                    if (!string.IsNullOrWhiteSpace(SerializedResults.Inputs.SortString))
                    {
                        returns.Where(p => p.Name.Equals("sort", StringComparison.OrdinalIgnoreCase)).First().Value = SerializedResults.Inputs.SortString;
                    }

                    if (SerializedResults.Inputs.StartRow.HasValue && SerializedResults.Inputs.StartRow.Value > -1)
                    {
                        returns.Where(p => p.Name.Equals("startrow", StringComparison.OrdinalIgnoreCase)).First().Value = SerializedResults.Inputs.StartRow.Value;
                    }

                    if (SerializedResults.Inputs.RowLimit.HasValue && SerializedResults.Inputs.RowLimit.Value > 0)
                    {
                        returns.Where(p => p.Name.Equals("rowlimit", StringComparison.OrdinalIgnoreCase)).First().Value = SerializedResults.Inputs.RowLimit.Value;
                    }

                    if (!string.IsNullOrWhiteSpace(SerializedResults.Inputs.Properties))
                    {
                        returns.Where(p => p.Name.Equals("properties", StringComparison.OrdinalIgnoreCase)).First().Value = SerializedResults.Inputs.Properties;
                    }

                    if (SerializedResults.Inputs.EnableStemming.HasValue)
                    {
                        returns.Where(p => p.Name.Equals("enablestemming", StringComparison.OrdinalIgnoreCase)).First().Value = SerializedResults.Inputs.EnableStemming.Value;
                    }

                    if (SerializedResults.Inputs.TrimDuplicates.HasValue)
                    {
                        returns.Where(p => p.Name.Equals("trimduplicates", StringComparison.OrdinalIgnoreCase)).First().Value = SerializedResults.Inputs.TrimDuplicates.Value;
                    }

                    if (SerializedResults.Inputs.EnableQueryRules.HasValue)
                    {
                        returns.Where(p => p.Name.Equals("enablequeryrules", StringComparison.OrdinalIgnoreCase)).First().Value = SerializedResults.Inputs.EnableQueryRules.Value;
                    }

                    if (SerializedResults.Inputs.ProcessBestBets.HasValue)
                    {
                        returns.Where(p => p.Name.Equals("processbestbets", StringComparison.OrdinalIgnoreCase)).First().Value = SerializedResults.Inputs.ProcessBestBets.Value;
                    }

                    if (SerializedResults.Inputs.ProcessPersonal.HasValue)
                    {
                        returns.Where(p => p.Name.Equals("processpersonal", StringComparison.OrdinalIgnoreCase)).First().Value = SerializedResults.Inputs.ProcessPersonal.Value;
                    }

                    if (SerializedResults.Inputs.EnableNicknames.HasValue)
                    {
                        returns.Where(p => p.Name.Equals("enablenicknames", StringComparison.OrdinalIgnoreCase)).First().Value = SerializedResults.Inputs.EnableNicknames.Value;
                    }

                    if (SerializedResults.Inputs.EnablePhonetic.HasValue)
                    {
                        returns.Where(p => p.Name.Equals("enablephonetic", StringComparison.OrdinalIgnoreCase)).First().Value = SerializedResults.Inputs.EnablePhonetic.Value;
                    }

                    if (SerializedResults.ExecutionTime.HasValue)
                    {
                        returns.Where(p => p.Name.Equals("executiontime", StringComparison.OrdinalIgnoreCase)).First().Value = SerializedResults.ExecutionTime.Value;
                    }

                    if (SerializedResults.ResultRows.HasValue)
                    {
                        returns.Where(p => p.Name.Equals("resultrows", StringComparison.OrdinalIgnoreCase)).First().Value = SerializedResults.ResultRows.Value;
                    }

                    if (SerializedResults.TotalRows.HasValue)
                    {
                        returns.Where(p => p.Name.Equals("totalrows", StringComparison.OrdinalIgnoreCase)).First().Value = SerializedResults.TotalRows.Value;
                    }

                    string resultsJson = JsonConvert.SerializeObject(SerializedResults);

                    returns.Where(p => p.Name.Equals("serializedresults", StringComparison.OrdinalIgnoreCase)).First().Value = resultsJson;

                    returns.Where(p => p.Name.Equals("responsestatus", StringComparison.OrdinalIgnoreCase)).First().Value = ResponseStatus.Success;
                }
                else
                {
                    throw new Exception("No results returned.");
                }
            }
            catch (Exception ex)
            {
                returns.Where(p => p.Name.Equals("responsestatus", StringComparison.OrdinalIgnoreCase)).First().Value = ResponseStatus.Error;
                returns.Where(p => p.Name.Equals("responsestatusdescription", StringComparison.OrdinalIgnoreCase)).First().Value = ex.Message;
            }
            serviceObject.Properties.BindPropertiesToResultTable();
        }


    }
}
