using K2Field.SmartObjects.Services.SharePoint.Search.Data;
using SourceCode.SharePoint15.Client;
using SourceCode.SmartObjects.Services.ServiceSDK;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace K2Field.SmartObjects.Services.SharePoint.Search.Utilities
{
    public static class BrokerUtils
    {
        public static SourceCode.SharePoint15.Client.Context InitializeContext(Configuration configuration)
        {
            SourceCode.SharePoint15.Client.Context result = null;
            try
            {
                if (configuration.Username.Length > 0)
                {
                    result = new Context(configuration.SiteUrl, configuration.Username, configuration.Password, configuration.Office365, configuration.ServiceTimeOut);
                }
                else
                {
                    if (!string.IsNullOrEmpty(configuration.OAuthToken))
                    {
                        result = new Context(configuration.SiteUrl, configuration.OAuthToken, configuration.ServiceTimeOut);
                    }
                    else
                    {
                        result = new Context(configuration.SiteUrl, configuration.ServiceTimeOut);
                    }
                }
            }
            catch (System.Exception ex)
            {
                //System.Text.StringBuilder stringBuilder = new System.Text.StringBuilder(SourceCode.SmartObjects.Services.SharePoint.Properties.Resources.error_ContextInitializationFailed);
                //stringBuilder.Append(System.Environment.NewLine);
                //stringBuilder.Append(string.Format(System.Globalization.CultureInfo.InvariantCulture, SourceCode.SmartObjects.Services.SharePoint.Properties.Resources.error_Url, new object[]
                //{
                //    siteUrl
                //}));
                //stringBuilder.Append(System.Environment.NewLine);
                //stringBuilder.Append(string.Format(System.Globalization.CultureInfo.InvariantCulture, SourceCode.SmartObjects.Services.SharePoint.Properties.Resources.error_Username, new object[]
                //{
                //    System.Security.Principal.WindowsIdentity.GetCurrent().Name
                //}));
                //stringBuilder.Append(System.Environment.NewLine);
                //stringBuilder.Append(System.Environment.NewLine);
                //stringBuilder.Append(SourceCode.SmartObjects.Services.SharePoint.Properties.Resources.error_ErrorDetails);
                //stringBuilder.Append(System.Environment.NewLine);
                //stringBuilder.Append(ex.Message);
                //stringBuilder.Append(System.Environment.NewLine);
                //stringBuilder.Append(string.Format(System.Globalization.CultureInfo.InvariantCulture, SourceCode.SmartObjects.Services.SharePoint.Properties.Resources.error_MethodName, new object[]
                //{
                //    "this.initializeContext"
                //}));
                
                //this.serviceBroker.ServicePackage.IsSuccessful = false;
                //throw new System.Exception(stringBuilder.ToString());
                throw ex;
            }
            return result;
        }

    }
}
