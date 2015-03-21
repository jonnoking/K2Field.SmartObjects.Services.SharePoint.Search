using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace K2Field.SmartObjects.Services.SharePoint.Search.Data
{
    public static class SPSearchSource
    {
        public static Dictionary<string, string> GetAllSourceIds()
        {
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

            SourceIds.Add("64cde128-76be-4943-b960-146e613a7e1e", "InternetSearchResults");
            SourceIds.Add("1dd9c4dc-8a6a-48a2-88b7-54dc3d97bf15", "InternetSearchSuggestions");
            SourceIds.Add("495318b6-0d9a-4d0f-939b-41cc17b49abd", "LocalPeopleSearchIndex");
            SourceIds.Add("5b557a96-b0ef-443c-8f55-fdcceb1e142a", "LocalSearchIndex");

            return SourceIds;
        }

        public static Dictionary<string, string> GetSourceIds()
        {
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

            return SourceIds;
        }

        public static Dictionary<string, string> GetOtherSourceIds()
        {
            Dictionary<string, string> SourceIds = new Dictionary<string, string>();
            SourceIds.Add("64cde128-76be-4943-b960-146e613a7e1e", "InternetSearchResults");
            SourceIds.Add("1dd9c4dc-8a6a-48a2-88b7-54dc3d97bf15", "InternetSearchSuggestions");
            SourceIds.Add("495318b6-0d9a-4d0f-939b-41cc17b49abd", "LocalPeopleSearchIndex");
            SourceIds.Add("5b557a96-b0ef-443c-8f55-fdcceb1e142a", "LocalSearchIndex");

            return SourceIds;
        }
    }
}
