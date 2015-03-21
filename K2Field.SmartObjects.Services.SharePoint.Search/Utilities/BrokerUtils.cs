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
using K2Field.SmartObjects.Services.SharePoint.Search.Data;

namespace K2Field.SmartObjects.Services.SharePoint.Search.Utilities
{
    public static class BrokerUtils
    {

        public static SearchInputs GetInputs(Property[] inputs)
        {
            SearchInputs InputValues = new SearchInputs();

            string search = string.Empty;
            var searchProp = inputs.Where(p => p.Name.Equals("search", StringComparison.OrdinalIgnoreCase)).FirstOrDefault();
            if (searchProp != null && searchProp.Value != null && !string.IsNullOrWhiteSpace(searchProp.Value.ToString()))
            {
                search = searchProp.Value.ToString();
                InputValues.Search = search;
            }
            else
            {
                throw new Exception("Search is a required property");
            }

            string searchsiteurl = string.Empty;
            var searchsiteurlprop = inputs.Where(p => p.Name.Equals("searchsiteurl", StringComparison.OrdinalIgnoreCase)).FirstOrDefault();
            if (searchsiteurlprop != null && searchsiteurlprop.Value != null && !string.IsNullOrWhiteSpace(searchsiteurlprop.Value.ToString()))
            {
                searchsiteurl = searchsiteurlprop.Value.ToString();
                InputValues.SiteUrl = searchsiteurl;
            }
            //else
            //{
            //    throw new Exception("Site Url is a required property");
            //}


            string properties = string.Empty;
            var propertiesprop = inputs.Where(p => p.Name.Equals("properties", StringComparison.OrdinalIgnoreCase)).FirstOrDefault();
            if (propertiesprop != null && propertiesprop.Value != null && !string.IsNullOrWhiteSpace(propertiesprop.Value.ToString()))
            {
                properties = propertiesprop.Value.ToString();
                InputValues.Properties = properties;
            }

            int startRow = -1;
            var startRowProp = inputs.Where(p => p.Name.Equals("startrow", StringComparison.OrdinalIgnoreCase)).FirstOrDefault();
            if (startRowProp != null && startRowProp.Value != null && !string.IsNullOrWhiteSpace(startRowProp.Value.ToString()))
            {
                if (int.TryParse(startRowProp.Value.ToString(), out startRow) && startRow > -1)
                {
                    InputValues.StartRow = startRow;
                }
            }

            int rowLimit = -1;
            var rowLimitProp = inputs.Where(p => p.Name.Equals("rowlimit", StringComparison.OrdinalIgnoreCase)).FirstOrDefault();
            if (rowLimitProp != null && rowLimitProp.Value != null && !string.IsNullOrWhiteSpace(rowLimitProp.Value.ToString()))
            {
                if (int.TryParse(rowLimitProp.Value.ToString(), out rowLimit) && rowLimit > 0)
                {
                    InputValues.RowLimit = rowLimit;
                }
            }

            Guid sourceid = Guid.Empty;
            var sourceidProp = inputs.Where(p => p.Name.Equals("sourceid", StringComparison.OrdinalIgnoreCase)).FirstOrDefault();
            if (sourceidProp != null && sourceidProp.Value != null && !string.IsNullOrWhiteSpace(sourceidProp.Value.ToString()))
            {
                if (Guid.TryParse(sourceidProp.Value.ToString(), out sourceid))
                {
                    InputValues.SourceId = sourceid;
                }
            }

            bool enablenicknames = false;
            var enablenicknamesProp = inputs.Where(p => p.Name.Equals("enablenicknames", StringComparison.OrdinalIgnoreCase)).FirstOrDefault();
            if (enablenicknamesProp != null && enablenicknamesProp.Value != null && !string.IsNullOrWhiteSpace(enablenicknamesProp.Value.ToString()))
            {
                if (bool.TryParse(enablenicknamesProp.Value.ToString(), out enablenicknames))
                {
                    InputValues.EnableNicknames = enablenicknames;
                }
            }

            bool enablephonetic = false;
            var enablephoneticProp = inputs.Where(p => p.Name.Equals("enablephonetic", StringComparison.OrdinalIgnoreCase)).FirstOrDefault();
            if (enablephoneticProp != null && enablephoneticProp.Value != null && !string.IsNullOrWhiteSpace(enablephoneticProp.Value.ToString()))
            {
                if (bool.TryParse(enablephoneticProp.Value.ToString(), out enablephonetic))
                {
                    InputValues.EnablePhonetic = enablephonetic;
                }
            }


            string fileext = string.Empty;
            var fileextprop = inputs.Where(p => p.Name.Equals("fileextensionsfilter", StringComparison.OrdinalIgnoreCase)).FirstOrDefault();
            if (fileextprop != null && fileextprop.Value != null && !string.IsNullOrWhiteSpace(fileextprop.Value.ToString()))
            {
                InputValues.FileExtensions = new List<string>();
                fileext = fileextprop.Value.ToString();
                string[] sortsArray = fileext.Split(',');

                foreach (string fx in sortsArray)
                {
                    InputValues.FileExtensions.Add(fx.Trim());
                }

                string filter = string.Empty;
                for (int i = 0; i < InputValues.FileExtensions.Count; i++)
                {
                    filter += "\"" + InputValues.FileExtensions[i] + "\"";
                    if (i <= InputValues.FileExtensions.Count - 2)
                    {
                        filter += ",";
                    }
                }
                InputValues.FileExtensionsString = filter;
            }


            string sorts = string.Empty;
            Dictionary<string, string> sort = new Dictionary<string, string>();
            var sortProp = inputs.Where(p => p.Name.Equals("sort", StringComparison.OrdinalIgnoreCase)).FirstOrDefault();
            if (sortProp != null && sortProp.Value != null && !string.IsNullOrWhiteSpace(sortProp.Value.ToString()))
            {
                sorts = sortProp.Value.ToString();
                string[] sortsArray = sorts.Split(',');
                foreach (string s in sortsArray)
                {
                    string[] ss = s.Split(':');
                    string prop = string.Empty;
                    string direction = string.Empty;
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
                                direction = "descending";
                                break;
                            case "ascending":
                            case "asc":
                                direction = "ascending";
                                break;
                            default:
                                direction = "ascending";
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

            if (sort.Count > 0)
            {
                string sortstring = string.Empty;
                InputValues.Sort = sort;

                int o = 0;
                foreach (KeyValuePair<string, string> s in sort)
                {
                    InputValues.SortString += s.Key + ":" + s.Value;
                    if (o <= sort.Count - 2)
                    {
                        InputValues.SortString += ",";
                    }
                    o++;
                }
            }

            bool enablestemming = false;
            var enablestemmingprop = inputs.Where(p => p.Name.Equals("enablestemming", StringComparison.OrdinalIgnoreCase)).FirstOrDefault();
            if (enablestemmingprop != null && enablestemmingprop.Value != null && !string.IsNullOrWhiteSpace(enablestemmingprop.Value.ToString()))
            {
                if (bool.TryParse(enablestemmingprop.Value.ToString(), out enablestemming))
                {
                    InputValues.EnableStemming = enablestemming;
                }
            }

            bool trimduplicates = false;
            var trimduplicatesprop = inputs.Where(p => p.Name.Equals("trimduplicates", StringComparison.OrdinalIgnoreCase)).FirstOrDefault();
            if (trimduplicatesprop != null && trimduplicatesprop.Value != null && !string.IsNullOrWhiteSpace(trimduplicatesprop.Value.ToString()))
            {
                if (bool.TryParse(trimduplicatesprop.Value.ToString(), out trimduplicates))
                {
                    InputValues.TrimDuplicates = trimduplicates;
                }
            }

            bool enablequeryrules = false;
            var enablequeryrulesprop = inputs.Where(p => p.Name.Equals("enablequeryrules", StringComparison.OrdinalIgnoreCase)).FirstOrDefault();
            if (enablequeryrulesprop != null && enablequeryrulesprop.Value != null && !string.IsNullOrWhiteSpace(enablequeryrulesprop.Value.ToString()))
            {
                if (bool.TryParse(enablequeryrulesprop.Value.ToString(), out enablequeryrules))
                {
                    InputValues.EnableQueryRules = enablequeryrules;
                }
            }

            bool processbestbets = false;
            var processbestbetsprop = inputs.Where(p => p.Name.Equals("processbestbets", StringComparison.OrdinalIgnoreCase)).FirstOrDefault();
            if (processbestbetsprop != null && processbestbetsprop.Value != null && !string.IsNullOrWhiteSpace(processbestbetsprop.Value.ToString()))
            {
                if (bool.TryParse(processbestbetsprop.Value.ToString(), out processbestbets))
                {
                    InputValues.ProcessBestBets = processbestbets;
                }
            }

            bool processpersonal = false;
            var processpersonalprop = inputs.Where(p => p.Name.Equals("processpersonal", StringComparison.OrdinalIgnoreCase)).FirstOrDefault();
            if (processpersonalprop != null && processpersonalprop.Value != null && !string.IsNullOrWhiteSpace(processpersonalprop.Value.ToString()))
            {
                if (bool.TryParse(processpersonalprop.Value.ToString(), out processpersonal))
                {
                    InputValues.ProcessPersonal = processpersonal;
                }
            }

            return InputValues;
        }


        //for debugging
        public static string GetColumns(RESTSearchResultsSerialized results)
        {
            string cols = string.Empty;
            int i = 0;
            foreach (ResultCell cell in results.SearchResults.Rows[0].Cells.OrderBy(p => p.Key))
            {
                cols += cell.Key + " (" + cell.ValueType + ")";
                if (i <= results.SearchResults.Rows.Count - 2)
                {
                    cols += ",";
                }
                i++;
            }

            return cols;
        }

        public static string BuildSearchText(SearchInputs Inputs, Configuration Configuration)
        {

            string RequestUri = Configuration.SiteUrl + "/_api/search/query";
            string SearchQuery = string.Empty;

            SearchQuery = "?querytext='" + Inputs.Search + "'";

            if (!string.IsNullOrWhiteSpace(Inputs.SiteUrl))
            {
                string p = "+path:\"" + Inputs.SiteUrl + "\"";
                SearchQuery = SearchQuery.Insert(SearchQuery.Length - 1, p);
            }

            SearchQuery += "&culture=" + Configuration.LocaleId;

            if (Inputs.StartRow.HasValue && Inputs.StartRow.Value > -1)
            {
                SearchQuery += "&startrow=" + Inputs.StartRow;
            }

            if (Inputs.RowLimit.HasValue && Inputs.RowLimit.Value > -1)
            {
                SearchQuery += "&rowlimit=" + Inputs.RowLimit;
            }

            if (Inputs.SourceId != null && Inputs.SourceId != Guid.Empty)
            {
                SearchQuery += "&sourceid='" + Inputs.SourceId + "'";
            }

            if (Inputs.Sort.Count > 0)
            {
                SearchQuery += "&sortlist='" + Inputs.SortString + "'";
            }

            if (!string.IsNullOrWhiteSpace(Inputs.Properties))
            {
                SearchQuery += "&Properties='" + Inputs.Properties + "'";
            }

            if (Inputs.FileExtensions != null && Inputs.FileExtensions.Count > 0)
            {
                //&refiners='filetype'

                if (Inputs.FileExtensions.Count < 2)
                {
                    SearchQuery += "&refiners='filetype,fileextension'&refinementfilters='filetype:equals(" + Inputs.FileExtensionsString + ")'";
                }
                else
                {
                    SearchQuery += "&refiners='filetype,fileextension'&refinementfilters='filetype:or(" + Inputs.FileExtensionsString + ")'";
                    //serviceBroker.ServicePackage.PageNumber
                    //serviceBroker.ServicePackage.PageSize;

                }
            }

            if (Inputs.EnableStemming.HasValue)
            {
                SearchQuery += "&enablestemming=" + Inputs.EnableStemming.ToString().ToLower();
            }

            if (Inputs.TrimDuplicates.HasValue)
            {
                SearchQuery += "&trimduplicates=" + Inputs.TrimDuplicates.ToString().ToLower();
            }

            if (Inputs.EnableQueryRules.HasValue)
            {
                SearchQuery += "&enablequeryrules=" + Inputs.EnableQueryRules.ToString().ToLower();
            }

            if (Inputs.ProcessBestBets.HasValue)
            {
                SearchQuery += "&processbestbets=" + Inputs.ProcessBestBets.ToString().ToLower();
            }

            if (Inputs.ProcessPersonal.HasValue)
            {
                SearchQuery += "&processpersonalfavorites=" + Inputs.ProcessPersonal.ToString().ToLower();
            }

            if (Inputs.EnableNicknames.HasValue)
            {
                SearchQuery += "&enablenicknames=" + Inputs.EnableNicknames.ToString().ToLower();
            }

            if (Inputs.EnablePhonetic.HasValue)
            {
                SearchQuery += "&enablephonetic=" + Inputs.EnablePhonetic.ToString().ToLower();
            }

            return RequestUri + SearchQuery;
        }


        public static RESTSearchResults ExecuteRESTRequest(string RequestUri, ServiceAssemblyBase serviceBroker)
        {
            var res = string.Empty;
            HttpWebRequest request = null;
            RESTSearchResults searchResults = null;
            
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
                    char[] sp = { '\\' };
                    string[] user = serviceBroker.Service.ServiceConfiguration.ServiceAuthentication.UserName.Split(sp);
                    if (user.Length > 1)
                    {
                        request.Credentials = new NetworkCredential(user[1], serviceBroker.Service.ServiceConfiguration.ServiceAuthentication.Password, user[0]);
                    }
                    else
                    {
                        request.Credentials = new NetworkCredential(serviceBroker.Service.ServiceConfiguration.ServiceAuthentication.UserName, serviceBroker.Service.ServiceConfiguration.ServiceAuthentication.Password);
                    }

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

        public static SearchInputs GetUserSearchInputs(Property[] inputs)
        {
            SearchInputs Inputs = GetInputs(inputs);
            Inputs.SourceId = Guid.Parse("b09a7990-05ea-4af9-81ef-edfab16c4e31");

            return Inputs;
        }

        public static RESTSearchResultsSerialized ExecuteSharePointSearch(Property[] inputs, RequiredProperties required, Configuration Configuration, ServiceAssemblyBase serviceBroker)
        {
            SearchInputs SearchInputs = Utilities.BrokerUtils.GetInputs(inputs);
            return ExecuteSearch(SearchInputs, Configuration, serviceBroker);
        }
        public static RESTSearchResultsSerialized ExecuteSharePointSearchRaw(Property[] inputs, RequiredProperties required, Configuration Configuration, ServiceAssemblyBase serviceBroker)
        {
            // Raw search = append input to end of querytext
            SearchInputs SearchInputs = new SearchInputs();
            string search = string.Empty;
            var searchProp = inputs.Where(p => p.Name.Equals("search", StringComparison.OrdinalIgnoreCase)).FirstOrDefault();
            if (searchProp != null && searchProp.Value != null && !string.IsNullOrWhiteSpace(searchProp.Value.ToString()))
            {
                search = searchProp.Value.ToString();
                SearchInputs.Search = search;
            }
            else
            {
                throw new Exception("Search is a required property");
            }

            string RequestUri = Configuration.SiteUrl + "/_api/search/query?querytext=" + search;

            return ProcessResults(ExecuteRESTRequest(RequestUri, serviceBroker), SearchInputs);
        }

        public static RESTSearchResultsSerialized ExecuteSharePointUserSearch(Property[] inputs, RequiredProperties required, Configuration Configuration, ServiceAssemblyBase serviceBroker)
        {
            SearchInputs SearchInputs = Utilities.BrokerUtils.GetUserSearchInputs(inputs);
            return ExecuteSearch(SearchInputs, Configuration, serviceBroker);
        }

        public static RESTSearchResultsSerialized ExecuteSearch(SearchInputs inputs, Configuration Configuration, ServiceAssemblyBase serviceBroker)
        {
            RESTSearchResults res = Utilities.BrokerUtils.ExecuteRESTRequest(Utilities.BrokerUtils.BuildSearchText(inputs, Configuration), serviceBroker);

            return ProcessResults(res, inputs);
        }

        public static RESTSearchResultsSerialized ProcessResults(RESTSearchResults res, SearchInputs inputs)
        {
            RESTSearchResultsSerialized SerializedResults = new RESTSearchResultsSerialized();
            SerializedResults.Inputs = inputs;
            
            if (res != null)
            {

                SerializedResults.ExecutionTime = res.ElapsedTime;

                if (res.PrimaryQueryResult != null && res.PrimaryQueryResult.RefinementResults != null)
                {
                    SerializedResults.TotalRows = res.PrimaryQueryResult.RelevantResults.TotalRows;

                    SerializedResults.ResultRows = res.PrimaryQueryResult.RelevantResults.RowCount;

                    SerializedResults.ResultTitle = res.PrimaryQueryResult.RelevantResults.ResultTitle;
                    SerializedResults.SearchResults = res.PrimaryQueryResult.RelevantResults.Table;
                    SerializedResults.ResultTitleUrl = res.PrimaryQueryResult.RelevantResults.ResultTitleUrl;
                }
                else
                {
                    SerializedResults.TotalRows = 0;

                    SerializedResults.ResultRows = 0;
                }

                SerializedResults.SpellingSuggestions = res.SpellingSuggestion;


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

        // old
        //public static SourceCode.SharePoint15.Client.Context InitializeContext(Configuration configuration)
        //{
        //    SourceCode.SharePoint15.Client.Context result = null;
        //    try
        //    {
        //        if (configuration.Username.Length > 0)
        //        {
        //            result = new Context(configuration.SiteUrl, configuration.Username, configuration.Password, configuration.Office365, configuration.ServiceTimeOut);
        //        }
        //        else
        //        {
        //            if (!string.IsNullOrEmpty(configuration.OAuthToken))
        //            {
        //                result = new Context(configuration.SiteUrl, configuration.OAuthToken, configuration.ServiceTimeOut);
        //            }
        //            else
        //            {
        //                result = new Context(configuration.SiteUrl, configuration.ServiceTimeOut);
        //            }
        //        }
        //    }
        //    catch (System.Exception ex)
        //    {
        //        //System.Text.StringBuilder stringBuilder = new System.Text.StringBuilder(SourceCode.SmartObjects.Services.SharePoint.Properties.Resources.error_ContextInitializationFailed);
        //        //stringBuilder.Append(System.Environment.NewLine);
        //        //stringBuilder.Append(string.Format(System.Globalization.CultureInfo.InvariantCulture, SourceCode.SmartObjects.Services.SharePoint.Properties.Resources.error_Url, new object[]
        //        //{
        //        //    siteUrl
        //        //}));
        //        //stringBuilder.Append(System.Environment.NewLine);
        //        //stringBuilder.Append(string.Format(System.Globalization.CultureInfo.InvariantCulture, SourceCode.SmartObjects.Services.SharePoint.Properties.Resources.error_Username, new object[]
        //        //{
        //        //    System.Security.Principal.WindowsIdentity.GetCurrent().Name
        //        //}));
        //        //stringBuilder.Append(System.Environment.NewLine);
        //        //stringBuilder.Append(System.Environment.NewLine);
        //        //stringBuilder.Append(SourceCode.SmartObjects.Services.SharePoint.Properties.Resources.error_ErrorDetails);
        //        //stringBuilder.Append(System.Environment.NewLine);
        //        //stringBuilder.Append(ex.Message);
        //        //stringBuilder.Append(System.Environment.NewLine);
        //        //stringBuilder.Append(string.Format(System.Globalization.CultureInfo.InvariantCulture, SourceCode.SmartObjects.Services.SharePoint.Properties.Resources.error_MethodName, new object[]
        //        //{
        //        //    "this.initializeContext"
        //        //}));
                
        //        //this.serviceBroker.ServicePackage.IsSuccessful = false;
        //        //throw new System.Exception(stringBuilder.ToString());
        //        throw ex;
        //    }
        //    return result;
        //}

    }
}
