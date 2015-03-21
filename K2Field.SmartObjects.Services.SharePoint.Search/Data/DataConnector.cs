using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using SourceCode.SmartObjects.Services.ServiceSDK;
using SourceCode.SmartObjects.Services.ServiceSDK.Objects;
using SourceCode.SmartObjects.Services.ServiceSDK.Types;

using K2Field.SmartObjects.Services.SharePoint.Search.Interfaces;
using SourceCode.Web.Utilities;
using System.Xml;
using K2Field.SmartObjects.Services.SharePoint.Search.Properties;
using System.Reflection;
using System.Globalization;

namespace K2Field.SmartObjects.Services.SharePoint.Search.Data
{
    class DataConnector : IDataConnector
    {
        #region Class Level Fields

        #region Constants
        /// <summary>
        /// Constant for the Type Mappings configuration lookup in the service instance.
        /// </summary>
        private static string __TypeMappings = "Type Mappings";
        #endregion

        #region Private Fields
        /// <summary>
        /// Local serviceBroker variable.
        /// </summary>
        private ServiceAssemblyBase serviceBroker = null;

        private string spsiteurl = string.Empty;
        private bool useOffice365 = false;

        private Configuration _configuration = new Configuration();
        public Configuration Configuration
        {
            get
            {
                return this._configuration;
            }
        }

        #endregion

        #endregion

        #region Constructor
        /// <summary>
        /// Instantiates a new DataConnector.
        /// </summary>
        /// <param name="serviceBroker">The ServiceBroker.</param>
        public DataConnector(ServiceAssemblyBase serviceBroker)
        {
            // Set local serviceBroker variable.
            this.serviceBroker = serviceBroker;
        }
        #endregion

        #region Methods

        #region void Dispose()
        /// <summary>
        /// Performs application-defined tasks associated with freeing, releasing, or resetting unmanaged resources.
        /// </summary>
        public void Dispose()
        {
            // Add any additional IDisposable implementation code here. Make sure to dispose of any data connections.
            // Clear references to serviceBroker.
            serviceBroker = null;
        }
        #endregion

        #region void GetConfiguration()
        /// <summary>
        /// Gets the configuration from the service instance and stores the retrieved configuration in local variables for later use.
        /// </summary>
        public void GetConfiguration()
        {
            this.Configuration.AdvancedSearchOptions = bool.Parse(this.GetServiceConfigurationValue("AdvancedSearchOptions", false));

            //string serviceConfigurationValue = this.GetServiceConfigurationValue(Resources.TermStoreId, false);
            this._configuration.SiteUrl = this.GetServiceConfigurationValue(Resources.label_SiteUrl, true);
            //this._configuration.DescribeType = ServiceObjectHelper.DescribeType.Full;
            if (this._configuration.SiteUrl.Contains("?describetype="))
            {
                string[] array = this._configuration.SiteUrl.Split(new string[]
		        {
			        "?describetype="
		        }, StringSplitOptions.RemoveEmptyEntries);
                this._configuration.SiteUrl = array[0];
                //        this._configuration.DescribeType = (ServiceObjectHelper.DescribeType)Enum.Parse(typeof(ServiceObjectHelper.DescribeType), array[1].Split(new char[]
                //{
                //    '&'
                //})[0], true);
                this.serviceBroker.Service.ServiceConfiguration["Siteurl"] = this._configuration.SiteUrl;
            }
            //this._configuration.DescribeSubsites = bool.Parse(this.GetServiceConfigurationValue(Resources.DescribeSubsites, false));
            //this._configuration.AdminSiteUrl = this.GetServiceConfigurationValue(Resources.AdminSiteUrl, false);
            this._configuration.Version = System.Diagnostics.FileVersionInfo.GetVersionInfo(Assembly.GetExecutingAssembly().Location).FileVersion;
            //this._configuration.Dynamic = bool.Parse(this.GetServiceConfigurationValue(Resources.Dynamic, false));
            this._configuration.Office365 = bool.Parse(this.GetServiceConfigurationValue(Resources.Office365, false));
            this._configuration.ServiceTimeOut = int.Parse(this.GetServiceConfigurationValue(Resources.ServiceTimeOut, true), CultureInfo.InvariantCulture);
            //this._configuration.IncludeHiddenLists = bool.Parse(this.GetServiceConfigurationValue(Resources.IncludeHiddenLists, false));
            //this._configuration.IncludeHiddenLibraries = bool.Parse(this.GetServiceConfigurationValue(Resources.IncludeHiddenLibraries, false));
            //this._configuration.DocumentsThrottle = int.Parse(this.GetServiceConfigurationValue(Resources.ThrottleDocuments, false), CultureInfo.InvariantCulture);
            //this._configuration.ListItemsThrottle = int.Parse(this.GetServiceConfigurationValue(Resources.ThrottleListItems, false), CultureInfo.InvariantCulture);
            //this._configuration.UseInternalFieldNames = bool.Parse(this.GetServiceConfigurationValue(Resources.UseInternalFieldNames, false));
            //this._configuration.ParseLookupFieldValues = bool.Parse(this.GetServiceConfigurationValue(Resources.ParseLookupFieldValues, false));
            //this._configuration.TermStoreId = (string.IsNullOrEmpty(serviceConfigurationValue) ? this._configuration.TermStoreId : new Guid(serviceConfigurationValue));
            this._configuration.LocaleId = int.Parse(this.GetServiceConfigurationValue(Resources.DefaultLocaleId, true), CultureInfo.InvariantCulture);
            this._configuration.Domain = "";
            this._configuration.Username = "";
            this._configuration.Password = null;
            this._configuration.OAuthToken = null;
            //this._configuration.IncludedFields = this.GetServiceConfigurationValue(Resources.IncludedFields, false);
            //this._configuration.ExcludedFields = this.GetServiceConfigurationValue(Resources.ExcludedFields, false);
            //this._configuration.IncludedLists = this.GetServiceConfigurationValue(Resources.IncludedLists, false);
            //this._configuration.ExcludedLists = this.GetServiceConfigurationValue(Resources.ExcludedLists, false);
            if (this.serviceBroker.Service.ServiceConfiguration.ServiceAuthentication.AuthenticationMode.Equals(AuthenticationMode.Static) || this.serviceBroker.Service.ServiceConfiguration.ServiceAuthentication.AuthenticationMode.Equals(AuthenticationMode.SSO))
            {
                this._configuration.Domain = this.serviceBroker.Service.ServiceConfiguration.ServiceAuthentication.Extra;
                if (this.serviceBroker.Service.ServiceConfiguration.ServiceAuthentication.UserName.Contains(":"))
                {
                    this._configuration.Username = this.serviceBroker.Service.ServiceConfiguration.ServiceAuthentication.UserName.Split(new char[]
			{
				':'
			})[1];
                }
                else
                {
                    this._configuration.Username = this.serviceBroker.Service.ServiceConfiguration.ServiceAuthentication.UserName;
                }
                this._configuration.Password = K2SecureString.FromString(this.serviceBroker.Service.ServiceConfiguration.ServiceAuthentication.Password);
                return;
            }
            if (this.serviceBroker.Service.ServiceConfiguration.ServiceAuthentication.AuthenticationMode.Equals(AuthenticationMode.OAuth))
            {
                this._configuration.OAuthToken = this.serviceBroker.Service.ServiceConfiguration.ServiceAuthentication.OAuthToken;
            }
        }
        #endregion

        #region void SetupConfiguration()
        /// <summary>
        /// Sets up the required configuration parameters in the service instance. When a new service instance is registered for this ServiceBroker, the configuration parameters are surfaced to the appropriate tooling. The configuration parameters are provided by the person registering the service instance.
        /// </summary>
        public void SetupConfiguration()
        {
            this.serviceBroker.Service.ServiceConfiguration.Add(Resources.label_SiteUrl, true, "http://");
            //this.serviceBroker.Service.ServiceConfiguration.Add(Resources.DescribeSubsites, false, "true");
            //this.serviceBroker.Service.ServiceConfiguration.Add(Resources.AdminSiteUrl, false, "");
            //this.serviceBroker.Service.ServiceConfiguration.Add(Resources.Dynamic, true, "false");
            this.serviceBroker.Service.ServiceConfiguration.Add(Resources.Office365, true, "false");
            this.serviceBroker.Service.ServiceConfiguration.Add(Resources.ServiceTimeOut, true, "120000");
            //this.serviceBroker.Service.ServiceConfiguration.Add(Resources.IncludeHiddenLists, false, "false");
            //this.serviceBroker.Service.ServiceConfiguration.Add(Resources.IncludeHiddenLibraries, false, "false");
            //this.serviceBroker.Service.ServiceConfiguration.Add(Resources.ThrottleDocuments, true, "50");
            //this.serviceBroker.Service.ServiceConfiguration.Add(Resources.ThrottleListItems, true, "2000");
            //this.serviceBroker.Service.ServiceConfiguration.Add(Resources.UseInternalFieldNames, false, "false");
            //this.serviceBroker.Service.ServiceConfiguration.Add(Resources.ParseLookupFieldValues, false, "true");
            //this.serviceBroker.Service.ServiceConfiguration.Add(Resources.TermStoreId, false, string.Empty);
            this.serviceBroker.Service.ServiceConfiguration.Add(Resources.DefaultLocaleId, false, CultureInfo.CurrentCulture.LCID.ToString(CultureInfo.InvariantCulture));
            //this.serviceBroker.Service.ServiceConfiguration.Add(Resources.IncludedFields, false, string.Empty);
            //this.serviceBroker.Service.ServiceConfiguration.Add(Resources.ExcludedFields, false, string.Empty);
            //this.serviceBroker.Service.ServiceConfiguration.Add(Resources.IncludedLists, false, string.Empty);
            //this.serviceBroker.Service.ServiceConfiguration.Add(Resources.ExcludedLists, false, string.Empty);

            this.serviceBroker.Service.ServiceConfiguration.Add("AdvancedSearchOptions", true, "false");
        }
        #endregion

        #region void SetupService()
        /// <summary>
        /// Sets up the service instance's default name, display name, and description.
        /// </summary>
        public void SetupService()
        {
            serviceBroker.Service.Name = "SharePoint2013Search" + Configuration.SiteUrl.Replace("https://", "").Replace("http://", "").Replace(".", "").Replace(":", "");
            serviceBroker.Service.MetaData.DisplayName = "SharePoint 2013 Search - " + Configuration.SiteUrl;
            serviceBroker.Service.MetaData.Description = "SharePoint 2013 Search - " + Configuration.SiteUrl;
        }
        #endregion

        #region void DescribeSchema()
        /// <summary>
        /// Describes the schema of the underlying data and services to the K2 platform.
        /// </summary>
        public void DescribeSchema()
        {
            TypeMappings map = GetTypeMappings();

            SPSearch search = new SPSearch(serviceBroker, Configuration);
            search.Create();
            SPSearchUser usersearch = new SPSearchUser(serviceBroker, Configuration);
            usersearch.Create();
            SPSearchDocument documentsearch = new SPSearchDocument(serviceBroker, Configuration);
            documentsearch.Create();
        }

        #endregion


        #region XmlDocument DiscoverSchema()
        /// <summary>
        /// Discovers the schema of the underlying data and services, and then maps the schema into a structure and format which is compliant with the requirements of Service Objects.
        /// </summary>
        /// <returns>An XmlDocument containing the discovered schema in a structure which complies with the requirements of Service Objects.</returns>
        public XmlDocument DiscoverSchema()
        {
            return null;
        }
        #endregion

        #region TypeMappings GetTypeMappings()
        /// <summary>
        /// Gets the type mappings used to map the underlying data's types to the appropriate K2 SmartObject types.
        /// </summary>
        /// <returns>A TypeMappings object containing the ServiceBroker's type mappings which were previously stored in the service instance configuration.</returns>
        public TypeMappings GetTypeMappings()
        {
            // Lookup and return the type mappings stored in the service instance.
            return (TypeMappings)serviceBroker.Service.ServiceConfiguration[__TypeMappings];
        }
        #endregion

        #region void SetTypeMappings()
        /// <summary>
        /// Sets the type mappings used to map the underlying data's types to the appropriate K2 SmartObject types.
        /// </summary>
        public void SetTypeMappings()
        {
            // Variable declaration.
            TypeMappings map = new TypeMappings();

            // Add type mappings.


            // Add the type mappings to the service instance.
            serviceBroker.Service.ServiceConfiguration.Add(__TypeMappings, map);
        }
        #endregion

        #region void Execute(Property[] inputs, RequiredProperties required, Property[] returns, MethodType methodType, ServiceObject serviceObject)
        /// <summary>
        /// Executes the Service Object method and returns any data.
        /// </summary>
        /// <param name="inputs">A Property[] array containing all the allowed input properties.</param>
        /// <param name="required">A RequiredProperties collection containing the required properties.</param>
        /// <param name="returns">A Property[] array containing all the allowed return properties.</param>
        /// <param name="methodType">A MethoType indicating what type of Service Object method was called.</param>
        /// <param name="serviceObject">A ServiceObject containing populated properties for use with the method call.</param>
        public void Execute(Property[] inputs, RequiredProperties required, Property[] returns, MethodType methodType, ServiceObject serviceObject)
        {
            this.GetConfiguration();

            if (serviceObject.Methods[0].Name.Equals("spsearch") || serviceObject.Methods[0].Name.Equals("spsearchraw") || serviceObject.Methods[0].Name.Equals("deserializesearchresults"))
            {
                SPSearch spsearch = new SPSearch(serviceBroker, this.Configuration);
                spsearch.ExecuteSearch(inputs, required, returns, methodType, serviceObject);
            }

            if (serviceObject.Methods[0].Name.Equals("spsearchread") || serviceObject.Methods[0].Name.Equals("spsearchrawread"))
            {
                SPSearch spsearch = new SPSearch(serviceBroker, this.Configuration);
                spsearch.ExecuteSearchRead(inputs, required, returns, methodType, serviceObject);
            }

            if (serviceObject.Methods[0].Name.Equals("listsourceidsstatic"))
            {
                SPSearch spsearch = new SPSearch(serviceBroker, this.Configuration);
                spsearch.ExecuteListSourceIds(inputs, required, returns, methodType, serviceObject);
            }

            if (serviceObject.Methods[0].Name.Equals("listothersourceidsstatic"))
            {
                SPSearch spsearch = new SPSearch(serviceBroker, this.Configuration);
                spsearch.ExecuteListOtherSourceIds(inputs, required, returns, methodType, serviceObject);
            }


            if (serviceObject.Methods[0].Name.Equals("spsearchusers") || serviceObject.Methods[0].Name.Equals("deserializeusersearchresults"))
            {
                SPSearchUser spsearchuser = new SPSearchUser(serviceBroker, this.Configuration);
                spsearchuser.ExecuteSearch(inputs, required, returns, methodType, serviceObject);
            }

            if (serviceObject.Methods[0].Name.Equals("spsearchusersread"))
            {
                SPSearchUser spsearchuser = new SPSearchUser(serviceBroker, this.Configuration);
                spsearchuser.ExecuteSearchRead(inputs, required, returns, methodType, serviceObject);
            }


            if (serviceObject.Methods[0].Name.Equals("spsearchdocuments") || serviceObject.Methods[0].Name.Equals("deserializedocumentsearchresults"))
            {
                SPSearchDocument spsearchdocumet = new SPSearchDocument(serviceBroker, this.Configuration);
                spsearchdocumet.ExecuteSearch(inputs, required, returns, methodType, serviceObject);
            }

            if (serviceObject.Methods[0].Name.Equals("spsearchdocumentsread"))
            {
                SPSearchDocument spsearchdocumet = new SPSearchDocument(serviceBroker, this.Configuration);
                spsearchdocumet.ExecuteSearchRead(inputs, required, returns, methodType, serviceObject);
            }
        }
        #endregion

        #endregion



        //

        public string GetServiceConfigurationValue(string serviceConfigObjectName, bool throwExceptionIfNull)
        {
            if (this.serviceBroker.Service.ServiceConfiguration[serviceConfigObjectName] != null)
            {
                return this.serviceBroker.Service.ServiceConfiguration[serviceConfigObjectName].ToString();
            }
            if (throwExceptionIfNull)
            {
                throw new ArgumentNullException(serviceConfigObjectName);
            }
            return string.Empty;
        }


    }


}
