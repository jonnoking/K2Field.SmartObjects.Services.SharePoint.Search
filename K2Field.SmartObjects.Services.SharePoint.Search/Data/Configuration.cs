using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace K2Field.SmartObjects.Services.SharePoint.Search.Data
{
    public class Configuration
    {
        public bool AdvancedSearchOptions
        { get; set; }
        public string AdminSiteUrl
        {
            get;
            set;
        }
        public string SiteUrl
        {
            get;
            set;
        }
        public string SiteTitle
        {
            get;
            set;
        }
        public System.Guid ListId
        {
            get;
            set;
        }
        public string ListRootFolderName
        {
            get;
            set;
        }
        public string ListTitle
        {
            get;
            set;
        }
        public string Version
        {
            get;
            set;
        }
        public bool DescribeSubsites
        {
            get;
            set;
        }
        public bool Dynamic
        {
            get;
            set;
        }
        public string Domain
        {
            get;
            set;
        }
        public string Username
        {
            get;
            set;
        }
        public System.Security.SecureString Password
        {
            get;
            set;
        }
        public string OAuthToken
        {
            get;
            set;
        }
        public bool Office365
        {
            get;
            set;
        }
        public int ServiceTimeOut
        {
            get;
            set;
        }
        public int ListItemsThrottle
        {
            get;
            set;
        }
        public int DocumentsThrottle
        {
            get;
            set;
        }
        public bool IncludeHiddenLists
        {
            get;
            set;
        }
        public bool IncludeHiddenListFields
        {
            get;
            set;
        }
        public bool IncludeHiddenLibraries
        {
            get;
            set;
        }
        public bool IncludeHiddenLibraryFields
        {
            get;
            set;
        }
        public bool UseInternalFieldNames
        {
            get;
            set;
        }
        public bool ParseLookupFieldValues
        {
            get;
            set;
        }
        //public ServiceObjectHelper.DescribeType DescribeType
        //{
        //    get;
        //    set;
        //}
        public System.Guid TermStoreId
        {
            get;
            set;
        }
        public int LocaleId
        {
            get;
            set;
        }
        public bool ExceptionOnEmtpyResult
        {
            get;
            set;
        }
        public string BindingName
        {
            get;
            set;
        }
        public string IncludedFields
        {
            get;
            set;
        }
        public string ExcludedFields
        {
            get;
            set;
        }
        public string IncludedLists
        {
            get;
            set;
        }
        public string ExcludedLists
        {
            get;
            set;
        }
    }

}
