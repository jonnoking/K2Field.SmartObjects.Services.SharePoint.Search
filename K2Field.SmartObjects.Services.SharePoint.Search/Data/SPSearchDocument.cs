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
    public class SPSearchDocument
    {
        private ServiceAssemblyBase serviceBroker = null;
        private Configuration Configuration { get; set; }

        public SPSearchDocument(ServiceAssemblyBase serviceBroker, Configuration configuration)
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
            SPSearchServiceObject.Name = "spsearchdocument";
            SPSearchServiceObject.MetaData.DisplayName = "SharePoint Document Search";

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

        private List<Property> InputsCore = new List<Property>();
        private List<Property> InputsAdvanced = new List<Property>();
        private List<Property> ReturnSummary = new List<Property>();
        private List<Property> ReturnSearch = new List<Property>();
        private List<Property> ReturnUserSearch = new List<Property>();
        private List<Property> ReturnStatus = new List<Property>();

        private List<Property> GetSPSearchProperties()
        {
            List<Property> ContainerProperties = new List<Property>();

            InputsCore = SPSearchProperties.GetCoreSearchInputProperties();
            ContainerProperties.AddRange(InputsCore);

            if (this.Configuration.AdvancedSearchOptions)
            {
                InputsAdvanced = SPSearchProperties.GetAdvancedSearchInputProperties();
                ContainerProperties.AddRange(InputsAdvanced);
            }

            ReturnSummary = SPSearchProperties.GetSearchResultSummaryProperties();
            ContainerProperties.AddRange(ReturnSummary);

            ReturnSearch = SPSearchProperties.GetContentSearchResultsProperties();
            ContainerProperties.AddRange(ReturnSearch);

            ReturnStatus = StandardReturns.GetStandardReturnProperties();
            ContainerProperties.AddRange(ReturnStatus);

            return ContainerProperties;
        }

        private Method CreateSearch(List<Property> SPSearchProps)
        {
            Method Search = new Method();
            Search.Name = "spsearchdocuments";
            Search.MetaData.DisplayName = "Search Documents";
            Search.Type = MethodType.List;

            Search.InputProperties.Add(SPSearchProps.Where(p => p.Name == "search").First());
            Search.Validation.RequiredProperties.Add(SPSearchProps.Where(p => p.Name == "search").First());
            Search.InputProperties.Add(SPSearchProps.Where(p => p.Name == "searchsiteurl").First());
            Search.InputProperties.Add(SPSearchProps.Where(p => p.Name == "fileextensionsfilter").First());

            Search.InputProperties.Add(SPSearchProps.Where(p => p.Name == "sort").First());
            Search.InputProperties.Add(SPSearchProps.Where(p => p.Name == "sourceid").First());

            Search.InputProperties.Add(SPSearchProps.Where(p => p.Name == "startrow").First());
            Search.InputProperties.Add(SPSearchProps.Where(p => p.Name == "rowlimit").First());

            Search.InputProperties.Add(SPSearchProps.Where(p => p.Name == "properties").First());


            // add inputs to return
            foreach (Property prop in InputsCore)
            {
                if (!Search.ReturnProperties.Contains(prop.Name))
                {
                    Search.ReturnProperties.Add(prop);
                }
            }

            // if advanced add to both inputs and returns
            if (this.Configuration.AdvancedSearchOptions)
            {
                foreach (Property prop in InputsAdvanced)
                {
                    if (!Search.InputProperties.Contains(prop.Name))
                    {
                        Search.InputProperties.Add(prop);
                    }
                    if (!Search.ReturnProperties.Contains(prop.Name))
                    {
                        Search.ReturnProperties.Add(prop);
                    }
                }
            }

            // add search summary to returns
            foreach (Property prop in ReturnSummary)
            {
                if (!Search.ReturnProperties.Contains(prop.Name))
                {
                    Search.ReturnProperties.Add(prop);
                }
            }

            // add search returns
            foreach (Property prop in ReturnSearch)
            {
                if (!Search.ReturnProperties.Contains(prop.Name))
                {
                    Search.ReturnProperties.Add(prop);
                }
            }

            // add status to returns
            foreach (Property prop in ReturnStatus)
            {
                if (!Search.ReturnProperties.Contains(prop.Name))
                {
                    Search.ReturnProperties.Add(prop);
                }
            }

            Search.ReturnProperties.Remove("serializedresults");
            Search.ReturnProperties.Remove("sourcename");

            return Search;
        }

        private Method CreateSearchRead(List<Property> SPSearchProps)
        {
            Method Search = new Method();
            Search.Name = "spsearchdocumentsread";
            Search.MetaData.DisplayName = "Search Documents Read";
            Search.Type = MethodType.Read;

            Search.InputProperties.Add(SPSearchProps.Where(p => p.Name == "search").First());
            Search.Validation.RequiredProperties.Add(SPSearchProps.Where(p => p.Name == "search").First());
            Search.InputProperties.Add(SPSearchProps.Where(p => p.Name == "searchsiteurl").First());
            Search.InputProperties.Add(SPSearchProps.Where(p => p.Name == "fileextensionsfilter").First());

            Search.InputProperties.Add(SPSearchProps.Where(p => p.Name == "sort").First());
            Search.InputProperties.Add(SPSearchProps.Where(p => p.Name == "sourceid").First());

            Search.InputProperties.Add(SPSearchProps.Where(p => p.Name == "startrow").First());
            Search.InputProperties.Add(SPSearchProps.Where(p => p.Name == "rowlimit").First());

            Search.InputProperties.Add(SPSearchProps.Where(p => p.Name == "properties").First());


            // add inputs to return
            foreach (Property prop in InputsCore)
            {
                if (!Search.ReturnProperties.Contains(prop.Name))
                {
                    Search.ReturnProperties.Add(prop);
                }
            }

            // if advanced add to both inputs and returns
            if (this.Configuration.AdvancedSearchOptions)
            {
                foreach (Property prop in InputsAdvanced)
                {
                    if (!Search.InputProperties.Contains(prop.Name))
                    {
                        Search.InputProperties.Add(prop);
                    }
                    if (!Search.ReturnProperties.Contains(prop.Name))
                    {
                        Search.ReturnProperties.Add(prop);
                    }
                }
            }

            // add search summary to returns -- includes serialized results
            foreach (Property prop in ReturnSummary)
            {
                if (!Search.ReturnProperties.Contains(prop.Name))
                {
                    Search.ReturnProperties.Add(prop);
                }
            }

            // add status to returns
            foreach (Property prop in ReturnStatus)
            {
                if (!Search.ReturnProperties.Contains(prop.Name))
                {
                    Search.ReturnProperties.Add(prop);
                }
            }

            Search.ReturnProperties.Remove("sourcename");

            return Search;
        }

        private Method CreateDeserializeSearchResults(List<Property> SPSearchProps)
        {
            Method Search = new Method();
            Search.Name = "deserializedocumentsearchresults";
            Search.MetaData.DisplayName = "Deserialize Document Search Results";
            Search.Type = MethodType.List;

            Search.InputProperties.Add(SPSearchProps.Where(p => p.Name == "serializedresults").First());

            // add inputs to return
            foreach (Property prop in InputsCore)
            {
                if (!Search.ReturnProperties.Contains(prop.Name))
                {
                    Search.ReturnProperties.Add(prop);
                }
            }

            // if advanced add to both inputs and returns
            if (this.Configuration.AdvancedSearchOptions)
            {
                foreach (Property prop in InputsAdvanced)
                {
                    if (!Search.InputProperties.Contains(prop.Name))
                    {
                        Search.InputProperties.Add(prop);
                    }
                    if (!Search.ReturnProperties.Contains(prop.Name))
                    {
                        Search.ReturnProperties.Add(prop);
                    }
                }
            }

            // add search summary to returns
            foreach (Property prop in ReturnSummary)
            {
                if (!Search.ReturnProperties.Contains(prop.Name))
                {
                    Search.ReturnProperties.Add(prop);
                }
            }

            // add search returns
            foreach (Property prop in ReturnSearch)
            {
                if (!Search.ReturnProperties.Contains(prop.Name))
                {
                    Search.ReturnProperties.Add(prop);
                }
            }

            // add status to returns
            foreach (Property prop in ReturnStatus)
            {
                if (!Search.ReturnProperties.Contains(prop.Name))
                {
                    Search.ReturnProperties.Add(prop);
                }
            }

            Search.ReturnProperties.Remove("serializedresults");
            Search.ReturnProperties.Remove("sourcename");

            return Search;
        }

        #endregion Describe


        #region Execute

        public void ExecuteSearch(Property[] inputs, RequiredProperties required, Property[] returns, MethodType methodType, ServiceObject serviceObject)
        {
            SPExecute Excute = new SPExecute(this.serviceBroker, this.Configuration);
            Excute.ExecuteSearch(inputs, required, returns, methodType, serviceObject);
        }

        public void ExecuteSearchRead(Property[] inputs, RequiredProperties required, Property[] returns, MethodType methodType, ServiceObject serviceObject)
        {
            SPExecute Excute = new SPExecute(this.serviceBroker, this.Configuration);
            Excute.ExecuteSearchRead(inputs, required, returns, methodType, serviceObject);
        }

        #endregion Execute

    }




}
