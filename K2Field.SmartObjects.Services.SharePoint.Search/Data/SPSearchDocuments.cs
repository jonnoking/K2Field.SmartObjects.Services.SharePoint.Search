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
    public class SPSearchDocuments
    {
        private ServiceAssemblyBase serviceBroker = null;
        private Configuration Configuration { get; set; }

        public SPSearchDocuments(ServiceAssemblyBase serviceBroker, Configuration configuration)
        {
            // Set local serviceBroker variable.
            this.serviceBroker = serviceBroker;
            this.Configuration = configuration;
        }

        #region Describe

        public void Create()
        {
            List<Property> SPSearchProps = GetSPSearchDocumentProperties();

            ServiceObject SPSearchServiceObject = new ServiceObject();
            SPSearchServiceObject.Name = "spsearchdocuments";
            SPSearchServiceObject.MetaData.DisplayName = "SharePoint Search Documents";

            SPSearchServiceObject.MetaData.ServiceProperties.Add("objecttype", "searchdocuments");

            SPSearchServiceObject.Active = true;

            foreach (Property prop in SPSearchProps)
            {
                if (!SPSearchServiceObject.Properties.Contains(prop.Name))
                {
                    SPSearchServiceObject.Properties.Add(prop);
                }
            }

            SPSearchServiceObject.Methods.Add(CreateSearchForDocuments(SPSearchProps));
            SPSearchServiceObject.Methods.Add(CreateSearchRead(SPSearchProps));
            SPSearchServiceObject.Methods.Add(CreateDeserializeSearchResults(SPSearchProps));


            serviceBroker.Service.ServiceObjects.Add(SPSearchServiceObject);
        }



        private List<Property> GetSPSearchDocumentProperties()
        {
            List<Property> ContainerProperties = new List<Property>();

            ContainerProperties.AddRange(SPSearchProperties.GetSearchInputProperties());
            ContainerProperties.AddRange(SPSearchProperties.GetSearchResultSummaryProperties());
            ContainerProperties.AddRange(SPSearchProperties.GetSearchResultReturnProperties());
            ContainerProperties.AddRange(StandardReturns.GetStandardReturnProperties());


            return ContainerProperties;
        }

        private Method CreateSearchForDocuments(List<Property> SPSearchProps)
        {
            Method Search = new Method();
            Search.Name = "spsearchfordocuments";
            Search.MetaData.DisplayName = "Search for Documents";
            Search.Type = MethodType.List;

            Search.InputProperties.Add(SPSearchProps.Where(p => p.Name == "search").First());
            Search.Validation.RequiredProperties.Add(SPSearchProps.Where(p => p.Name == "search").First());
            Search.InputProperties.Add(SPSearchProps.Where(p => p.Name == "fileextensions").First());

            Search.InputProperties.Add(SPSearchProps.Where(p => p.Name == "startrow").First());
            Search.InputProperties.Add(SPSearchProps.Where(p => p.Name == "rowlimit").First());
            Search.InputProperties.Add(SPSearchProps.Where(p => p.Name == "sort").First());
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
            Search.Name = "spsearchfordocuments";
            Search.MetaData.DisplayName = "Search for Documents Read";
            Search.Type = MethodType.Read;

            Search.InputProperties.Add(SPSearchProps.Where(p => p.Name == "search").First());
            Search.Validation.RequiredProperties.Add(SPSearchProps.Where(p => p.Name == "search").First());
            Search.InputProperties.Add(SPSearchProps.Where(p => p.Name == "fileextensions").First());

            Search.InputProperties.Add(SPSearchProps.Where(p => p.Name == "startrow").First());
            Search.InputProperties.Add(SPSearchProps.Where(p => p.Name == "rowlimit").First());
            Search.InputProperties.Add(SPSearchProps.Where(p => p.Name == "sort").First());
            Search.InputProperties.Add(SPSearchProps.Where(p => p.Name == "enablenicknames").First());
            Search.InputProperties.Add(SPSearchProps.Where(p => p.Name == "enablephonetic").First());


            Search.ReturnProperties.Add(SPSearchProps.Where(p => p.Name == "search").First());
            Search.ReturnProperties.Add(SPSearchProps.Where(p => p.Name == "startrow").First());
            Search.ReturnProperties.Add(SPSearchProps.Where(p => p.Name == "rowlimit").First());
            Search.ReturnProperties.Add(SPSearchProps.Where(p => p.Name == "sort").First());
            Search.ReturnProperties.Add(SPSearchProps.Where(p => p.Name == "enablenicknames").First());
            Search.ReturnProperties.Add(SPSearchProps.Where(p => p.Name == "enablephonetic").First());

            Search.ReturnProperties.Add(SPSearchProps.Where(p => p.Name == "executiontime").First());
            Search.ReturnProperties.Add(SPSearchProps.Where(p => p.Name == "resultrows").First());
            Search.ReturnProperties.Add(SPSearchProps.Where(p => p.Name == "totalrows").First());
            Search.ReturnProperties.Add(SPSearchProps.Where(p => p.Name == "tabletype").First());
            Search.ReturnProperties.Add(SPSearchProps.Where(p => p.Name == "resulttitle").First());
            Search.ReturnProperties.Add(SPSearchProps.Where(p => p.Name == "resulttitleurl").First());
            Search.ReturnProperties.Add(SPSearchProps.Where(p => p.Name == "spellingsuggestions").First());
            Search.ReturnProperties.Add(SPSearchProps.Where(p => p.Name == "serializedresults").First());
            Search.ReturnProperties.Add(SPSearchProps.Where(p => p.Name == "responsestatus").First());
            Search.ReturnProperties.Add(SPSearchProps.Where(p => p.Name == "responsestatusdescription").First());

            return Search;
        }

        private Method CreateDeserializeSearchResults(List<Property> SPSearchProps)
        {
            Method Search = new Method();
            Search.Name = "deserializesearchfordocumentsresults";
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
                    SerializedResults = ExecuteSharePointSearch(inputs, required, returns, methodType, serviceObject);
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
                        dr["resulttitle"] = SerializedResults.ResultTitle;
                        dr["resulttitleurl"] = SerializedResults.ResultTitleUrl;
                        dr["tabletype"] = SerializedResults.TableType;
                        dr["spellingsuggestions"] = SerializedResults.SpellingSuggestions;


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

                SerializedResults = ExecuteSharePointSearch(inputs, required, returns, methodType, serviceObject);

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

                    if (SerializedResults.Inputs.EnableNicknames.HasValue && SerializedResults.Inputs.EnableNicknames.Value)
                    {
                        returns.Where(p => p.Name.Equals("enablenicknames", StringComparison.OrdinalIgnoreCase)).First().Value = SerializedResults.Inputs.EnableNicknames.Value;
                    }

                    if (SerializedResults.Inputs.EnablePhonetic.HasValue && SerializedResults.Inputs.EnablePhonetic.Value)
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

                    returns.Where(p => p.Name.Equals("resulttitle", StringComparison.OrdinalIgnoreCase)).First().Value = SerializedResults.ResultTitle;
                    returns.Where(p => p.Name.Equals("resulttitleurl", StringComparison.OrdinalIgnoreCase)).First().Value = SerializedResults.ResultTitleUrl;
                    returns.Where(p => p.Name.Equals("tabletype", StringComparison.OrdinalIgnoreCase)).First().Value = SerializedResults.TableType;
                    returns.Where(p => p.Name.Equals("spellingsuggestions", StringComparison.OrdinalIgnoreCase)).First().Value = SerializedResults.SpellingSuggestions;

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

        // deprecated
        public void ExecuteDeserializeSearchResults(Property[] inputs, RequiredProperties required, Property[] returns, MethodType methodType, ServiceObject serviceObject)
        {
            serviceObject.Properties.InitResultTable();
            System.Data.DataRow dr;

            string json = string.Empty;

            try
            {
                Property SerializedProp = inputs.Where(p => p.Name.Equals("serializedresults", StringComparison.OrdinalIgnoreCase)).First();
                if (SerializedProp != null && SerializedProp.Value != null)
                {
                    json = SerializedProp.Value.ToString();
                }

                //IEnumerable<IDictionary<string, object>> searchResults = JsonConvert.DeserializeObject<IEnumerable<IDictionary<string, object>>>(json.Trim());

                SearchResultsSerialized SerializedSearch = JsonConvert.DeserializeObject<SearchResultsSerialized>(json.Trim());

                if (string.IsNullOrWhiteSpace(json) || SerializedSearch == null)
                {
                    throw new Exception("Failed to deserialize search results");
                }

                foreach (IDictionary<string, object> result in SerializedSearch.SearchResults)
                {
                    dr = serviceBroker.ServicePackage.ResultTable.NewRow();

                    if (!string.IsNullOrWhiteSpace(SerializedSearch.Inputs.Search))
                    {
                        dr["search"] = SerializedSearch.Inputs.Search;
                    }

                    if (SerializedSearch.Inputs.StartRow.HasValue && SerializedSearch.Inputs.StartRow.Value > -1)
                    {
                        dr["startrow"] = SerializedSearch.Inputs.StartRow.Value;
                    }

                    if (SerializedSearch.Inputs.RowLimit.HasValue && SerializedSearch.Inputs.RowLimit.Value > 0)
                    {
                        dr["rowlimit"] = SerializedSearch.Inputs.RowLimit.Value;
                    }

                    if (!string.IsNullOrWhiteSpace(SerializedSearch.Inputs.SortString))
                    {
                        dr["sort"] = SerializedSearch.Inputs.SortString;
                    }

                    if (!string.IsNullOrWhiteSpace(SerializedSearch.Inputs.SortString))
                    {
                        dr["sourceid"] = SerializedSearch.Inputs.SourceId;
                    }

                    if (SerializedSearch.ExecutionTime.HasValue)
                    {
                        dr["executiontime"] = SerializedSearch.ExecutionTime;
                    }

                    if (SerializedSearch.ResultRows.HasValue)
                    {
                        dr["resultrows"] = SerializedSearch.ResultRows;
                    }
                    if (SerializedSearch.TotalRows.HasValue)
                    {
                        dr["totalrows"] = SerializedSearch.TotalRows;
                    }
                    dr["resulttitle"] = SerializedSearch.ResultTitle;
                    dr["resulttitleurl"] = SerializedSearch.ResultTitleUrl;
                    dr["tabletype"] = SerializedSearch.TableType;
                    dr["spellingsuggestions"] = SerializedSearch.SpellingSuggestions;

                    foreach (string s in result.Keys)
                    {
                        if (result[s] != null)
                        {
                            dr[s.ToLower()] = result[s];
                        }
                    }
                    dr["responsestatus"] = ResponseStatus.Success;
                    serviceBroker.ServicePackage.ResultTable.Rows.Add(dr);
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


        public RESTSearchResultsSerialized ExecuteSharePointSearch(Property[] inputs, RequiredProperties required, Property[] returns, MethodType methodType, ServiceObject serviceObject)
        {
            serviceObject.Properties.InitResultTable();

            ClientResult<ResultTableCollection> results = null;

            SearchInputs SearchInputs = GetInputs(inputs);

            RESTSearchResultsSerialized SerializedResults = new RESTSearchResultsSerialized();
            SerializedResults.Inputs = SearchInputs;



            //KeywordQuery keywordQuery = new KeywordQuery(cc);
            //keywordQuery.QueryText = SearchInputs.Search;

            //SearchExecutor searchExecutor = new SearchExecutor(cc);
            if (SearchInputs.StartRow.HasValue && SearchInputs.StartRow.Value > -1)
            {
                //keywordQuery.StartRow = SearchInputs.StartRow.Value;
            }

            //keywordQuery.RowsPerPage = int.Parse(txtRowsPerPage.Text);
            if (SearchInputs.RowLimit.HasValue && SearchInputs.RowLimit.Value > -1)
            {
                //keywordQuery.RowLimit = SearchInputs.RowLimit.Value;
            }

            //keywordQuery.Culture = Configuration.LocaleId;

            if (SearchInputs.SourceId != null && SearchInputs.SourceId != Guid.Empty)
            {
                //keywordQuery.SourceId = SearchInputs.SourceId;
            }

            if (SearchInputs.Sort.Count > 0)
            {
                
            }

            if (SearchInputs.EnableNicknames.HasValue && SearchInputs.EnableNicknames.Value)
            {
                //keywordQuery.EnableNicknames = SearchInputs.EnableNicknames.Value;
            }

            if (SearchInputs.EnablePhonetic.HasValue && SearchInputs.EnablePhonetic.Value)
            {
                //keywordQuery.EnablePhonetic = SearchInputs.EnablePhonetic.Value;
            }


            // updated for inputs
            RESTSearchResults res = ExecuteRESTRequest(BuildSearchText(SearchInputs));


            if (res != null)
            {

                int executiontime = res.ElapsedTime;

                int totalresults = res.PrimaryQueryResult.RelevantResults.TotalRows;

                int resultrows = res.PrimaryQueryResult.RelevantResults.RowCount;


                SerializedResults.ResultTitle = res.PrimaryQueryResult.RelevantResults.ResultTitle;
                SerializedResults.ResultTitleUrl = res.PrimaryQueryResult.RelevantResults.ResultTitleUrl;
                SerializedResults.SpellingSuggestions = res.SpellingSuggestion;

                SerializedResults.SearchResults = res.PrimaryQueryResult.RelevantResults.Table;


                // set SourceId from execution results
                Guid sid = Guid.Empty;


                SearchProperty SourceId = res.Properties.Where(p => p.Key.Equals("sourceid", StringComparison.InvariantCultureIgnoreCase)).First();
                if (SourceId != null && Guid.TryParse(SourceId.Value, out sid))
                {
                    SerializedResults.Inputs.SourceId = sid;
                }



            }

            return SerializedResults;
        }


        public SearchInputs GetInputs(Property[] inputs)
        {
            SearchInputs InputValues = new SearchInputs();

            string search = string.Empty;
            var searchProp = inputs.Where(p => p.Name.Equals("search", StringComparison.OrdinalIgnoreCase)).First();
            if (searchProp != null && searchProp.Value != null && !string.IsNullOrWhiteSpace(searchProp.Value.ToString()))
            {
                search = searchProp.Value.ToString();
                InputValues.Search = search;
            }
            else
            {
                throw new Exception("Search is a required property");
            }

            int startRow = -1;
            var startRowProp = inputs.Where(p => p.Name.Equals("startrow", StringComparison.OrdinalIgnoreCase)).First();
            if (startRowProp != null && startRowProp.Value != null && !string.IsNullOrWhiteSpace(startRowProp.Value.ToString()))
            {
                if (int.TryParse(startRowProp.Value.ToString(), out startRow) && startRow > -1)
                {
                    InputValues.StartRow = startRow;
                }
            }

            int rowLimit = -1;
            var rowLimitProp = inputs.Where(p => p.Name.Equals("rowlimit", StringComparison.OrdinalIgnoreCase)).First();
            if (rowLimitProp != null && rowLimitProp.Value != null && !string.IsNullOrWhiteSpace(rowLimitProp.Value.ToString()))
            {
                if (int.TryParse(rowLimitProp.Value.ToString(), out rowLimit) && rowLimit > 0)
                {
                    InputValues.RowLimit = rowLimit;
                }
            }

            Guid sourceid = Guid.Empty;
            var sourceidProp = inputs.Where(p => p.Name.Equals("sourceid", StringComparison.OrdinalIgnoreCase)).First();
            if (sourceidProp != null && sourceidProp.Value != null && !string.IsNullOrWhiteSpace(sourceidProp.Value.ToString()))
            {
                if (Guid.TryParse(sourceidProp.Value.ToString(), out sourceid))
                {
                    InputValues.SourceId = sourceid;
                }
            }

            bool enablenicknames = false;
            var enablenicknamesProp = inputs.Where(p => p.Name.Equals("enablenicknames", StringComparison.OrdinalIgnoreCase)).First();
            if (enablenicknamesProp != null && enablenicknamesProp.Value != null && !string.IsNullOrWhiteSpace(enablenicknamesProp.Value.ToString()))
            {
                if (bool.TryParse(enablenicknamesProp.Value.ToString(), out enablenicknames))
                {
                    InputValues.EnableNicknames = enablenicknames;
                }
            }

            bool enablephonetic = false;
            var enablephoneticProp = inputs.Where(p => p.Name.Equals("enablephonetic", StringComparison.OrdinalIgnoreCase)).First();
            if (enablephoneticProp != null && enablephoneticProp.Value != null && !string.IsNullOrWhiteSpace(enablephoneticProp.Value.ToString()))
            {
                if (bool.TryParse(enablephoneticProp.Value.ToString(), out enablephonetic))
                {
                    InputValues.EnablePhonetic = enablephonetic;
                }
            }

            string sorts = string.Empty;
            Dictionary<string, SortDirection> sort = new Dictionary<string, SortDirection>();
            var sortProp = inputs.Where(p => p.Name.Equals("sort", StringComparison.OrdinalIgnoreCase)).First();
            if (sortProp != null && sortProp.Value != null && !string.IsNullOrWhiteSpace(sortProp.Value.ToString()))
            {
                sorts = sortProp.Value.ToString();
                string[] sortsArray = sorts.Split(';');
                foreach (string s in sortsArray)
                {
                    string[] ss = s.Split(',');
                    string prop = string.Empty;
                    Microsoft.SharePoint.Client.Search.Query.SortDirection direction;
                    if (ss.Length > 1)
                    {
                        // JJK: can we check if the supplied property exists?
                        prop = ss[0].Trim();
                        string dir = ss[1].Trim();
                        switch (dir.ToLower())
                        {
                            case "descending":
                            case "desc":
                            case "des":
                                direction = SortDirection.Descending;
                                break;
                            case "ascending":
                            case "asc":
                                direction = SortDirection.Ascending;
                                break;
                            default:
                                direction = SortDirection.Ascending;
                                break;
                        }

                        if (!string.IsNullOrWhiteSpace(prop))
                        {
                            sort.Add(prop, direction);
                        }
                    }
                }
                //returns.Where(p => p.Name.Equals("sort", StringComparison.OrdinalIgnoreCase)).First().Value = sorts;
            }
            InputValues.SortString = sorts;
            if (sort.Count > 0)
            {
                //InputValues.Sort = sort;
            }

            return InputValues;
        }


        private string BuildSearchText(SearchInputs Inputs)
        {

            string RequestUri = Configuration.SiteUrl + "/_api/search/query";
            string SearchQuery = string.Empty;

            SearchQuery = "?querytext='" + Inputs.Search + "'";

            if (Inputs.StartRow.HasValue && Inputs.StartRow.Value > -1)
            {
                SearchQuery += "&startrow=" + Inputs.StartRow;
            }

            //keywordQuery.RowsPerPage = int.Parse(txtRowsPerPage.Text);
            if (Inputs.RowLimit.HasValue && Inputs.RowLimit.Value > -1)
            {
                SearchQuery += "&rowlimit=" + Inputs.RowLimit;
            }

            //keywordQuery.Culture = Configuration.LocaleId;

            if (Inputs.SourceId != null && Inputs.SourceId != Guid.Empty)
            {
                SearchQuery += "&sourceid='" + Inputs.SourceId + "'";
            }

            if (Inputs.Sort.Count > 0)
            {
                
            }

            if (Inputs.EnableNicknames.HasValue && Inputs.EnableNicknames.Value)
            {
                SearchQuery += "&enablenicknames=" + Inputs.EnableNicknames;
            }

            if (Inputs.EnablePhonetic.HasValue && Inputs.EnablePhonetic.Value)
            {
                SearchQuery += "&enablephonetic=" + Inputs.EnablePhonetic;
            }

            return RequestUri + SearchQuery;
        }


        private RESTSearchResults ExecuteRESTRequest(string RequestUri)
        {
            var res = string.Empty;
            HttpWebRequest request = null;
            RESTSearchResults searchResults = null;
            //List<T> items = new List<T>();

            //string accessToken = Configuration.OAuthToken;

            try
            {
                request = (HttpWebRequest)WebRequest.Create(RequestUri);
                request.Method = "GET";
                request.Accept = "application/json";
                //                request.Expect = "100-continue";
                request.Headers.Add("Accept-Encoding", "gzip, deflate");

                if (serviceBroker.Service.ServiceConfiguration.ServiceAuthentication.AuthenticationMode == AuthenticationMode.Impersonate || serviceBroker.Service.ServiceConfiguration.ServiceAuthentication.AuthenticationMode == AuthenticationMode.ServiceAccount)
                {
                    request.UseDefaultCredentials = true;
                }
                if (serviceBroker.Service.ServiceConfiguration.ServiceAuthentication.AuthenticationMode == AuthenticationMode.OAuth)
                {
                    string accessToken = serviceBroker.Service.ServiceConfiguration.ServiceAuthentication.OAuthToken;
                    string headerBearer = String.Format(CultureInfo.InvariantCulture, "Bearer {0}", accessToken);

                    request.Headers.Add("Authorization", headerBearer.ToString());
                }
                if (serviceBroker.Service.ServiceConfiguration.ServiceAuthentication.AuthenticationMode == AuthenticationMode.Static)
                {
                    request.Credentials = GetCredentials();//unlikely to work for office 365
                }


                request.AutomaticDecompression = DecompressionMethods.GZip | DecompressionMethods.Deflate;

                using (HttpWebResponse Response = (HttpWebResponse)request.GetResponse())
                {
                    using (Stream st = Response.GetResponseStream())
                    {
                        using (StreamReader sr = new StreamReader(st))
                        {
                            res = sr.ReadToEnd();
                        }

                        searchResults = Newtonsoft.Json.JsonConvert.DeserializeObject<RESTSearchResults>(res);
                    }
                }
            }
            catch (WebException wex)
            {
                // should throw exception to force reauth
                throw;
            }
            catch (Exception ex)
            {
                throw;
            }
            finally
            {
                request = null;
            }
            return searchResults;
        }

        private ICredentials GetCredentials()
        {
            char[] sp = { '\\' };
            string[] user = serviceBroker.Service.ServiceConfiguration.ServiceAuthentication.UserName.Split(sp);
            if (user.Length > 1)
            {
                return new NetworkCredential(user[1], serviceBroker.Service.ServiceConfiguration.ServiceAuthentication.Password, user[0]);
            }
            else
            {
                return new NetworkCredential(serviceBroker.Service.ServiceConfiguration.ServiceAuthentication.UserName, serviceBroker.Service.ServiceConfiguration.ServiceAuthentication.Password);
            }
        }

        #endregion Execute

    }




}
