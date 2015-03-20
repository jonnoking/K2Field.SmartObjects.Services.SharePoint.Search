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
    public class SPSearchUser
    {
        private ServiceAssemblyBase serviceBroker = null;
        private Configuration Configuration { get; set; }

        public SPSearchUser(ServiceAssemblyBase serviceBroker, Configuration configuration)
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
            SPSearchServiceObject.Name = "spsearchuser";
            SPSearchServiceObject.MetaData.DisplayName = "SharePoint Search User";

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

            serviceBroker.Service.ServiceObjects.Add(SPSearchServiceObject);
        }

        private List<Property> GetSPSearchProperties()
        {
            List<Property> ContainerProperties = new List<Property>();

            ContainerProperties.AddRange(SPSearchProperties.GetSearchInputProperties());
            ContainerProperties.AddRange(SPSearchProperties.GetSearchResultSummaryProperties());
            ContainerProperties.AddRange(SPSearchProperties.GetStandardSearchReturnProperties());
            ContainerProperties.AddRange(SPSearchProperties.GetUserSearchResultProperties());
            ContainerProperties.AddRange(StandardReturns.GetStandardReturnProperties());

            return ContainerProperties;
        }

        private Method CreateSearch(List<Property> SPSearchProps)
        {
            Method Search = new Method();
            Search.Name = "spsearchusers";
            Search.MetaData.DisplayName = "Search Users";
            Search.Type = MethodType.List;

            Search.InputProperties.Add(SPSearchProps.Where(p => p.Name == "search").First());
            Search.Validation.RequiredProperties.Add(SPSearchProps.Where(p => p.Name == "search").First());

            Search.InputProperties.Add(SPSearchProps.Where(p => p.Name == "sort").First());

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

            Search.ReturnProperties.Remove("enablestemming");
            Search.ReturnProperties.Remove("trimduplicates");
            Search.ReturnProperties.Remove("enablequeryrules");
            Search.ReturnProperties.Remove("processbestbets");
            Search.ReturnProperties.Remove("processpersonal");

            Search.ReturnProperties.Remove("enablenicknames");
            Search.ReturnProperties.Remove("enablephonetic");


            return Search;
        }

        private Method CreateSearchRead(List<Property> SPSearchProps)
        {
            Method Search = new Method();
            Search.Name = "spsearchusersread";
            Search.MetaData.DisplayName = "Search Users Read";
            Search.Type = MethodType.Read;

            Search.InputProperties.Add(SPSearchProps.Where(p => p.Name == "search").First());
            Search.Validation.RequiredProperties.Add(SPSearchProps.Where(p => p.Name == "search").First());
            Search.InputProperties.Add(SPSearchProps.Where(p => p.Name == "sort").First());

            Search.InputProperties.Add(SPSearchProps.Where(p => p.Name == "startrow").First());
            Search.InputProperties.Add(SPSearchProps.Where(p => p.Name == "rowlimit").First());


            Search.ReturnProperties.Add(SPSearchProps.Where(p => p.Name == "search").First());
            Search.ReturnProperties.Add(SPSearchProps.Where(p => p.Name == "sort").First());
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

                Search.ReturnProperties.Remove("enablestemming");
                Search.ReturnProperties.Remove("trimduplicates");
                Search.ReturnProperties.Remove("enablequeryrules");
                Search.ReturnProperties.Remove("processbestbets");
                Search.ReturnProperties.Remove("processpersonal");

                Search.ReturnProperties.Remove("enablenicknames");
                Search.ReturnProperties.Remove("enablephonetic");

            }

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
                var sps = inputs.Where(p => p.Name.Equals("serializedresults", StringComparison.OrdinalIgnoreCase));
                if (sps.Count() > 0 && inputs.Where(p => p.Name.Equals("serializedresults", StringComparison.OrdinalIgnoreCase)).First() != null && inputs.Where(p => p.Name.Equals("serializedresults", StringComparison.OrdinalIgnoreCase)).First().Value != null)
                {
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
                else
                {
                    // if Search
                    SerializedResults = Utilities.BrokerUtils.ExecuteSharePointUserSearch(inputs, required, Configuration, serviceBroker);
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

                        if (SerializedResults.Inputs.StartRow.HasValue && SerializedResults.Inputs.StartRow.Value > -1)
                        {
                            dr["startrow"] = SerializedResults.Inputs.StartRow.Value;
                        }

                        if (SerializedResults.Inputs.RowLimit.HasValue && SerializedResults.Inputs.RowLimit.Value > 0)
                        {
                            dr["rowlimit"] = SerializedResults.Inputs.RowLimit.Value;
                        }

                        if (SerializedResults.Inputs.SourceId != null && SerializedResults.Inputs.SourceId != Guid.Empty)
                        {
                            dr["sourceid"] = SerializedResults.Inputs.SourceId;
                        }

                        if (!string.IsNullOrWhiteSpace(SerializedResults.Inputs.SortString))
                        {
                            dr["sort"] = SerializedResults.Inputs.SortString;
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
                RESTSearchResultsSerialized SerializedResults = new RESTSearchResultsSerialized();

                SerializedResults = Utilities.BrokerUtils.ExecuteSharePointUserSearch(inputs, required, Configuration, serviceBroker);

                if (SerializedResults != null)
                {
                    if (!string.IsNullOrWhiteSpace(SerializedResults.Inputs.Search))
                    {
                        returns.Where(p => p.Name.Equals("search", StringComparison.OrdinalIgnoreCase)).First().Value = SerializedResults.Inputs.Search;
                    }

                    if (SerializedResults.Inputs.StartRow.HasValue && SerializedResults.Inputs.StartRow.Value > -1)
                    {
                        returns.Where(p => p.Name.Equals("startrow", StringComparison.OrdinalIgnoreCase)).First().Value = SerializedResults.Inputs.StartRow.Value;
                    }

                    if (SerializedResults.Inputs.RowLimit.HasValue && SerializedResults.Inputs.RowLimit.Value > 0)
                    {
                        returns.Where(p => p.Name.Equals("rowlimit", StringComparison.OrdinalIgnoreCase)).First().Value = SerializedResults.Inputs.RowLimit.Value;
                    }

                    if (SerializedResults.Inputs.SourceId != null && SerializedResults.Inputs.SourceId != Guid.Empty)
                    {
                        returns.Where(p => p.Name.Equals("sourceid", StringComparison.OrdinalIgnoreCase)).First().Value = SerializedResults.Inputs.SourceId;
                    }

                    if (!string.IsNullOrWhiteSpace(SerializedResults.Inputs.SortString))
                    {
                        returns.Where(p => p.Name.Equals("sort", StringComparison.OrdinalIgnoreCase)).First().Value = SerializedResults.Inputs.SortString;
                    }
                    // for testing
                    //returns.Where(p => p.Name.Equals("sort", StringComparison.OrdinalIgnoreCase)).First().Value = Utilities.BrokerUtils.GetColumns(SerializedResults);

                    if (!string.IsNullOrWhiteSpace(SerializedResults.Inputs.FileExtensionsString))
                    {
                        returns.Where(p => p.Name.Equals("fileextensions", StringComparison.OrdinalIgnoreCase)).First().Value = SerializedResults.Inputs.FileExtensionsString;
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

      
        #endregion Execute

    }




}
