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
    public class SPSearch
    {
        private ServiceAssemblyBase serviceBroker = null;
        private Configuration Configuration { get; set; }

        public SPSearch(ServiceAssemblyBase serviceBroker, Configuration configuration)
        {
            // Set local serviceBroker variable.
            this.serviceBroker = serviceBroker;
            this.Configuration = configuration;
        }

        #region Describe

        public void Create()
        {
            //List<Property> SPSearchProps = GetSPSearchProperties();

            List<Property> SPSearchProps = GetSPSearchProperties();

            ServiceObject SPSearchServiceObject = new ServiceObject();
            SPSearchServiceObject.Name = "spsearch";
            SPSearchServiceObject.MetaData.DisplayName = "SharePoint Search";

            SPSearchServiceObject.MetaData.ServiceProperties.Add("objecttype", "search");

            SPSearchServiceObject.Active = true;

            foreach (Property prop in SPSearchProps)
            {
                if (!SPSearchServiceObject.Properties.Contains(prop.Name))
                {
                    SPSearchServiceObject.Properties.Add(prop);
                }
            }

            SPSearchServiceObject.Methods.Add(CreateSearch(SPSearchProps));
            SPSearchServiceObject.Methods.Add(CreateSearchRead(SPSearchProps));
            SPSearchServiceObject.Methods.Add(CreateDeserializeSearchResults(SPSearchProps));
            SPSearchServiceObject.Methods.Add(CreateSearchRaw(SPSearchProps));
            SPSearchServiceObject.Methods.Add(CreateSearchRawRead(SPSearchProps));
            SPSearchServiceObject.Methods.Add(CreateListSourceIds(SPSearchProps));
            SPSearchServiceObject.Methods.Add(CreateListOtherSourceIds(SPSearchProps));

            serviceBroker.Service.ServiceObjects.Add(SPSearchServiceObject);
        }

        private List<Property> GetSPSearchProperties()
        {
            List<Property> ContainerProperties = new List<Property>();

            ContainerProperties.AddRange(SPSearchProperties.GetSearchInputProperties());
            ContainerProperties.AddRange(SPSearchProperties.GetSearchResultSummaryProperties());
            ContainerProperties.AddRange(SPSearchProperties.GetSearchResultsProperties());
            ContainerProperties.AddRange(SPSearchProperties.GetUserSearchResultProperties());
            ContainerProperties.AddRange(StandardReturns.GetStandardReturnProperties());

            return ContainerProperties;
        }

        private Method CreateSearch(List<Property> SPSearchProps)
        {
            Method Search = new Method();
            Search.Name = "spsearch";
            Search.MetaData.DisplayName = "Search";
            Search.Type = MethodType.List;

            Search.InputProperties.Add(SPSearchProps.Where(p => p.Name == "search").First());
            Search.Validation.RequiredProperties.Add(SPSearchProps.Where(p => p.Name == "search").First());
            Search.InputProperties.Add(SPSearchProps.Where(p => p.Name == "searchsiteurl").First());
            Search.InputProperties.Add(SPSearchProps.Where(p => p.Name == "fileextensions").First());

            Search.InputProperties.Add(SPSearchProps.Where(p => p.Name == "sort").First());
            Search.InputProperties.Add(SPSearchProps.Where(p => p.Name == "sourceid").First());

            Search.InputProperties.Add(SPSearchProps.Where(p => p.Name == "startrow").First());
            Search.InputProperties.Add(SPSearchProps.Where(p => p.Name == "rowlimit").First());

            Search.InputProperties.Add(SPSearchProps.Where(p => p.Name == "properties").First());

            Search.InputProperties.Add(SPSearchProps.Where(p => p.Name == "enablestemming").First());
            Search.InputProperties.Add(SPSearchProps.Where(p => p.Name == "trimduplicates").First());
            Search.InputProperties.Add(SPSearchProps.Where(p => p.Name == "enablequeryrules").First());
            Search.InputProperties.Add(SPSearchProps.Where(p => p.Name == "processbestbets").First());
            Search.InputProperties.Add(SPSearchProps.Where(p => p.Name == "processpersonal").First());
            
            Search.InputProperties.Add(SPSearchProps.Where(p => p.Name == "enablenicknames").First());
            Search.InputProperties.Add(SPSearchProps.Where(p => p.Name == "enablephonetic").First());

            foreach (Property prop in SPSearchProps)
            {
                if (!Search.ReturnProperties.Contains(prop.Name))
                {
                    Search.ReturnProperties.Add(prop);
                }
            }
            if (Search.ReturnProperties.Contains("serializedresults"))
            {
                Search.ReturnProperties.Remove("serializedresults");
            }

            return Search;
        }

        private Method CreateSearchRead(List<Property> SPSearchProps)
        {
            Method Search = new Method();
            Search.Name = "spsearchread";
            Search.MetaData.DisplayName = "Search Read";
            Search.Type = MethodType.Read;

            Search.InputProperties.Add(SPSearchProps.Where(p => p.Name == "search").First());
            Search.Validation.RequiredProperties.Add(SPSearchProps.Where(p => p.Name == "search").First());
            Search.InputProperties.Add(SPSearchProps.Where(p => p.Name == "searchsiteurl").First());
            Search.InputProperties.Add(SPSearchProps.Where(p => p.Name == "fileextensions").First());

            Search.InputProperties.Add(SPSearchProps.Where(p => p.Name == "sort").First());
            Search.InputProperties.Add(SPSearchProps.Where(p => p.Name == "sourceid").First());

            Search.InputProperties.Add(SPSearchProps.Where(p => p.Name == "startrow").First());
            Search.InputProperties.Add(SPSearchProps.Where(p => p.Name == "rowlimit").First());

            Search.InputProperties.Add(SPSearchProps.Where(p => p.Name == "properties").First());

            Search.InputProperties.Add(SPSearchProps.Where(p => p.Name == "enablestemming").First());
            Search.InputProperties.Add(SPSearchProps.Where(p => p.Name == "trimduplicates").First());
            Search.InputProperties.Add(SPSearchProps.Where(p => p.Name == "enablequeryrules").First());
            Search.InputProperties.Add(SPSearchProps.Where(p => p.Name == "processbestbets").First());
            Search.InputProperties.Add(SPSearchProps.Where(p => p.Name == "processpersonal").First());

            Search.InputProperties.Add(SPSearchProps.Where(p => p.Name == "enablenicknames").First());
            Search.InputProperties.Add(SPSearchProps.Where(p => p.Name == "enablephonetic").First());


            Search.ReturnProperties.Add(SPSearchProps.Where(p => p.Name == "search").First());
            Search.ReturnProperties.Add(SPSearchProps.Where(p => p.Name == "searchsiteurl").First());
            Search.ReturnProperties.Add(SPSearchProps.Where(p => p.Name == "fileextensions").First());
            Search.ReturnProperties.Add(SPSearchProps.Where(p => p.Name == "sort").First());
            Search.ReturnProperties.Add(SPSearchProps.Where(p => p.Name == "sourceid").First());
            Search.ReturnProperties.Add(SPSearchProps.Where(p => p.Name == "startrow").First());
            Search.ReturnProperties.Add(SPSearchProps.Where(p => p.Name == "rowlimit").First());

            Search.ReturnProperties.Add(SPSearchProps.Where(p => p.Name == "properties").First());
            Search.ReturnProperties.Add(SPSearchProps.Where(p => p.Name == "enablestemming").First());
            Search.ReturnProperties.Add(SPSearchProps.Where(p => p.Name == "trimduplicates").First());
            Search.ReturnProperties.Add(SPSearchProps.Where(p => p.Name == "enablequeryrules").First());
            Search.ReturnProperties.Add(SPSearchProps.Where(p => p.Name == "processbestbets").First());
            Search.ReturnProperties.Add(SPSearchProps.Where(p => p.Name == "processpersonal").First());

            Search.ReturnProperties.Add(SPSearchProps.Where(p => p.Name == "enablenicknames").First());
            Search.ReturnProperties.Add(SPSearchProps.Where(p => p.Name == "enablephonetic").First());

            Search.ReturnProperties.Add(SPSearchProps.Where(p => p.Name == "executiontime").First());
            Search.ReturnProperties.Add(SPSearchProps.Where(p => p.Name == "resultrows").First());
            Search.ReturnProperties.Add(SPSearchProps.Where(p => p.Name == "totalrows").First());
            //Search.ReturnProperties.Add(SPSearchProps.Where(p => p.Name == "tabletype").First());
            //Search.ReturnProperties.Add(SPSearchProps.Where(p => p.Name == "resulttitle").First());
            //Search.ReturnProperties.Add(SPSearchProps.Where(p => p.Name == "resulttitleurl").First());
            //Search.ReturnProperties.Add(SPSearchProps.Where(p => p.Name == "spellingsuggestions").First());
            Search.ReturnProperties.Add(SPSearchProps.Where(p => p.Name == "serializedresults").First());
            Search.ReturnProperties.Add(SPSearchProps.Where(p => p.Name == "responsestatus").First());
            Search.ReturnProperties.Add(SPSearchProps.Where(p => p.Name == "responsestatusdescription").First());

            return Search;
        }

        private Method CreateSearchRaw(List<Property> SPSearchProps)
        {
            Method Search = new Method();
            Search.Name = "spsearchraw";
            Search.MetaData.DisplayName = "Search Raw";
            Search.Type = MethodType.List;

            Search.InputProperties.Add(SPSearchProps.Where(p => p.Name == "search").First());
            Search.Validation.RequiredProperties.Add(SPSearchProps.Where(p => p.Name == "search").First());

            Search.InputProperties.Add(SPSearchProps.Where(p => p.Name == "startrow").First());
            Search.InputProperties.Add(SPSearchProps.Where(p => p.Name == "rowlimit").First()); 

            foreach (Property prop in SPSearchProps)
            {
                if (!Search.ReturnProperties.Contains(prop.Name))
                {
                    Search.ReturnProperties.Add(prop);
                }
            }
            if (Search.ReturnProperties.Contains("serializedresults"))
            {
                Search.ReturnProperties.Remove("serializedresults");
            }

            return Search;
        }

        private Method CreateSearchRawRead(List<Property> SPSearchProps)
        {
            Method Search = new Method();
            Search.Name = "spsearchrawread";
            Search.MetaData.DisplayName = "Search Raw Read";
            Search.Type = MethodType.Read;

            Search.InputProperties.Add(SPSearchProps.Where(p => p.Name == "search").First());
            Search.Validation.RequiredProperties.Add(SPSearchProps.Where(p => p.Name == "search").First());

            Search.InputProperties.Add(SPSearchProps.Where(p => p.Name == "startrow").First());
            Search.InputProperties.Add(SPSearchProps.Where(p => p.Name == "rowlimit").First()); 

            Search.ReturnProperties.Add(SPSearchProps.Where(p => p.Name == "search").First());
            Search.ReturnProperties.Add(SPSearchProps.Where(p => p.Name == "startrow").First());
            Search.ReturnProperties.Add(SPSearchProps.Where(p => p.Name == "rowlimit").First()); 

            Search.ReturnProperties.Add(SPSearchProps.Where(p => p.Name == "executiontime").First());
            Search.ReturnProperties.Add(SPSearchProps.Where(p => p.Name == "resultrows").First());
            Search.ReturnProperties.Add(SPSearchProps.Where(p => p.Name == "totalrows").First());
            //Search.ReturnProperties.Add(SPSearchProps.Where(p => p.Name == "tabletype").First());
            //Search.ReturnProperties.Add(SPSearchProps.Where(p => p.Name == "resulttitle").First());
            //Search.ReturnProperties.Add(SPSearchProps.Where(p => p.Name == "resulttitleurl").First());
            //Search.ReturnProperties.Add(SPSearchProps.Where(p => p.Name == "spellingsuggestions").First());
            Search.ReturnProperties.Add(SPSearchProps.Where(p => p.Name == "serializedresults").First());
            Search.ReturnProperties.Add(SPSearchProps.Where(p => p.Name == "responsestatus").First());
            Search.ReturnProperties.Add(SPSearchProps.Where(p => p.Name == "responsestatusdescription").First());

            return Search;
        }


        private Method CreateDeserializeSearchResults(List<Property> SPSearchProps)
        {
            Method Search = new Method();
            Search.Name = "deserializesearchresults";
            Search.MetaData.DisplayName = "Deserialize Search Results";
            Search.Type = MethodType.List;

            Search.InputProperties.Add(SPSearchProps.Where(p => p.Name == "serializedresults").First());

            foreach (Property prop in SPSearchProps)
            {
                if (!Search.ReturnProperties.Contains(prop.Name))
                {
                    Search.ReturnProperties.Add(prop);
                }
            }
            if (Search.ReturnProperties.Contains("serializedresults"))
            {
                Search.ReturnProperties.Remove("serializedresults");
            }

            return Search;
        }

        private Method CreateListSourceIds(List<Property> SPSearchProps)
        {
            Method Search = new Method();
            Search.Name = "listsourceidsstatic";
            Search.MetaData.DisplayName = "List Source Ids Static";
            Search.Type = MethodType.List;

            Search.ReturnProperties.Add(SPSearchProps.Where(p => p.Name == "sourceid").First());
            Search.ReturnProperties.Add(SPSearchProps.Where(p => p.Name == "sourcename").First());

            return Search;
        }

        private Method CreateListOtherSourceIds(List<Property> SPSearchProps)
        {
            Method Search = new Method();
            Search.Name = "listothersourceidsstatic";
            Search.MetaData.DisplayName = "List Other Source Ids Static";
            Search.Type = MethodType.List;

            Search.ReturnProperties.Add(SPSearchProps.Where(p => p.Name == "sourceid").First());
            Search.ReturnProperties.Add(SPSearchProps.Where(p => p.Name == "sourcename").First());

            return Search;
        }


        #endregion Describe


        #region Execute

        public void ExecuteSearch(Property[] inputs, RequiredProperties required, Property[] returns, MethodType methodType, ServiceObject serviceObject)
        {
            serviceObject.Properties.InitResultTable();
            System.Data.DataRow dr;
            try
            {
                RESTSearchResultsSerialized SerializedResults = null;

                // if deserializesearchresults
                if(serviceObject.Methods[0].Name.Equals("deserializesearchresults"))
                { 
                //var sps = inputs.Where(p => p.Name.Equals("serializedresults", StringComparison.OrdinalIgnoreCase));
                //if (sps.Count() > 0 && inputs.Where(p => p.Name.Equals("serializedresults", StringComparison.OrdinalIgnoreCase)).First() != null && inputs.Where(p => p.Name.Equals("serializedresults", StringComparison.OrdinalIgnoreCase)).First().Value != null)
                //{
                    Property SerializedProp = inputs.Where(p => p.Name.Equals("serializedresults", StringComparison.OrdinalIgnoreCase)).First();
                    //if (SerializedProp != null && SerializedProp.Value != null)
                    //{
                    string json = string.Empty;
                    json = SerializedProp.Value.ToString();

                    //IEnumerable<IDictionary<string, object>> searchResults = JsonConvert.DeserializeObject<IEnumerable<IDictionary<string, object>>>(json.Trim());

                    SerializedResults = JsonConvert.DeserializeObject<RESTSearchResultsSerialized>(json.Trim());

                    if (string.IsNullOrWhiteSpace(json) || SerializedResults == null)
                    {
                        throw new Exception("Failed to deserialize search results");
                    }
                    //}
                }

                if (serviceObject.Methods[0].Name.Equals("spsearch"))
                {
                    // if Search
                    SerializedResults = Utilities.BrokerUtils.ExecuteSharePointSearch(inputs, required, Configuration, serviceBroker);
                }

                if (serviceObject.Methods[0].Name.Equals("spsearchraw"))
                {
                    // if Search Raw Read
                    SerializedResults = Utilities.BrokerUtils.ExecuteSharePointSearchRaw(inputs, required, Configuration, serviceBroker);
                }


                if (SerializedResults != null)
                {
                    // needs updating for REST
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
                            dr["fileextensions"] = SerializedResults.Inputs.FileExtensionsString;
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
                        //dr["resulttitle"] = SerializedResults.ResultTitle;
                        //dr["resulttitleurl"] = SerializedResults.ResultTitleUrl;
                        //dr["tabletype"] = SerializedResults.TableType;
                        //dr["spellingsuggestions"] = SerializedResults.SpellingSuggestions;


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
                //SearchInputs SearchInputs = GetInputs(inputs);
                RESTSearchResultsSerialized SerializedResults = new RESTSearchResultsSerialized();

                if (serviceObject.Methods[0].Name.Equals("spsearchread"))
                {
                    SerializedResults = Utilities.BrokerUtils.ExecuteSharePointSearch(inputs, required, Configuration, serviceBroker);
                }

                if (serviceObject.Methods[0].Name.Equals("spsearchrawread"))
                {
                    SerializedResults = Utilities.BrokerUtils.ExecuteSharePointSearchRaw(inputs, required, Configuration, serviceBroker);
                }

                if (SerializedResults != null)
                {
                    if (!string.IsNullOrWhiteSpace(SerializedResults.Inputs.Search))
                    {
                        returns.Where(p => p.Name.Equals("search", StringComparison.OrdinalIgnoreCase)).First().Value = SerializedResults.Inputs.Search;
                    }
                    
                    if (!string.IsNullOrWhiteSpace(SerializedResults.Inputs.SiteUrl))
                    {
                        returns.Where(p => p.Name.Equals("searchsiteurl", StringComparison.OrdinalIgnoreCase)).First().Value = SerializedResults.Inputs.Search;
                    }
                    
                    if (!string.IsNullOrWhiteSpace(SerializedResults.Inputs.FileExtensionsString))
                    {
                        returns.Where(p => p.Name.Equals("fileextensions", StringComparison.OrdinalIgnoreCase)).First().Value = SerializedResults.Inputs.FileExtensionsString;
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
                    // for testing
                    //returns.Where(p => p.Name.Equals("sort", StringComparison.OrdinalIgnoreCase)).First().Value = Utilities.BrokerUtils.GetColumns(SerializedResults);

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

                    //returns.Where(p => p.Name.Equals("resulttitle", StringComparison.OrdinalIgnoreCase)).First().Value = SerializedResults.ResultTitle;
                    //returns.Where(p => p.Name.Equals("resulttitleurl", StringComparison.OrdinalIgnoreCase)).First().Value = SerializedResults.ResultTitleUrl;
                    //returns.Where(p => p.Name.Equals("tabletype", StringComparison.OrdinalIgnoreCase)).First().Value = SerializedResults.TableType;
                    //returns.Where(p => p.Name.Equals("spellingsuggestions", StringComparison.OrdinalIgnoreCase)).First().Value = SerializedResults.SpellingSuggestions;

                    //string resultsJson = JsonConvert.SerializeObject(results.Value[0].ResultRows);
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


        public void ExecuteListSourceIds(Property[] inputs, RequiredProperties required, Property[] returns, MethodType methodType, ServiceObject serviceObject)
        {
            serviceObject.Properties.InitResultTable();
            System.Data.DataRow dr;

            Dictionary<string, string> SourceIds = new Dictionary<string, string>();
            SourceIds.Add("8413cd39-2156-4e00-b54d-11efd9abdb89", "Local SharePoint Results");
            SourceIds.Add("b09a7990-05ea-4af9-81ef-edfab16c4e31", "Local People Results");
            SourceIds.Add("203fba36-2763-4060-9931-911ac8c0583b", "Local Reports And Data Results");
            SourceIds.Add("78b793ce-7956-4669-aa3b-451fc5defebf", "Local Video Results");
            SourceIds.Add("e7ec8cee-ded8-43c9-beb5-436b54b31e84", "Documents");
            SourceIds.Add("5dc9f503-801e-4ced-8a2c-5d1237132419", "Items matching a content type");
            SourceIds.Add("e1327b9c-2b8c-4b23-99c9-3730cb29c3f7", "Items matching a tag");
            SourceIds.Add("48fec42e-4a92-48ce-8363-c2703a40e67d", "Items related to current user");
            SourceIds.Add("5c069288-1d17-454a-8ac6-9c642a065f48", "Items with same keyword as this item");
            SourceIds.Add("5e34578e-4d08-4edc-8bf3-002acf3cdbcc", "Pages");
            SourceIds.Add("38403c8c-3975-41a8-826e-717f2d41568a", "Pictures");
            SourceIds.Add("97c71db1-58ce-4891-8b64-585bc2326c12", "Popular");
            SourceIds.Add("ba63bbae-fa9c-42c0-b027-9a878f16557c", "Recently changed items");
            SourceIds.Add("ec675252-14fa-4fbe-84dd-8d098ed74181", "Recommended Items");
            SourceIds.Add("9479bf85-e257-4318-b5a8-81a180f5faa1", "Wiki");

            foreach (var Source in SourceIds)
            {
                dr = serviceBroker.ServicePackage.ResultTable.NewRow();
                Guid sid = Guid.Empty;
                if (Guid.TryParse(Source.Key, out sid))
                {
                    dr["sourceid"] = sid;
                    dr["sourcename"] = Source.Value;
                }
                serviceBroker.ServicePackage.ResultTable.Rows.Add(dr);
            }

        }

        public void ExecuteListOtherSourceIds(Property[] inputs, RequiredProperties required, Property[] returns, MethodType methodType, ServiceObject serviceObject)
        {
            serviceObject.Properties.InitResultTable();
            System.Data.DataRow dr;

            Dictionary<string, string> SourceIds = new Dictionary<string, string>();
            SourceIds.Add("64cde128-76be-4943-b960-146e613a7e1e", "InternetSearchResults");
            SourceIds.Add("1dd9c4dc-8a6a-48a2-88b7-54dc3d97bf15", "InternetSearchSuggestions");
            SourceIds.Add("495318b6-0d9a-4d0f-939b-41cc17b49abd", "LocalPeopleSearchIndex");
            SourceIds.Add("5b557a96-b0ef-443c-8f55-fdcceb1e142a", "LocalSearchIndex");

            foreach (var Source in SourceIds)
            {
                dr = serviceBroker.ServicePackage.ResultTable.NewRow();
                Guid sid = Guid.Empty;
                if (Guid.TryParse(Source.Key, out sid))
                {
                    dr["sourceid"] = sid;
                    dr["sourcename"] = Source.Value;
                }
                serviceBroker.ServicePackage.ResultTable.Rows.Add(dr);
            }

        }
        
        #endregion Execute

    }




}
