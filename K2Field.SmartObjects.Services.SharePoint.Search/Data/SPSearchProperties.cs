using SourceCode.SmartObjects.Services.ServiceSDK.Objects;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace K2Field.SmartObjects.Services.SharePoint.Search.Data
{
    public class SPSearchProperties
    {

        public static List<Property> GetSearchInputProperties()
        {
            List<Property> ContainerProperties = new List<Property>();

            Property search = new Property
            {
                Name = "search",
                MetaData = new MetaData("Search", "Search"),
                SoType = SourceCode.SmartObjects.Services.ServiceSDK.Types.SoType.Text,
                Type = "System.String"
            };
            ContainerProperties.Add(search);

            Property searchsite = new Property
            {
                Name = "searchsiteurl",
                MetaData = new MetaData("Search Site Url", "Search Site Url"),
                SoType = SourceCode.SmartObjects.Services.ServiceSDK.Types.SoType.Text,
                Type = "System.String"
            };
            ContainerProperties.Add(searchsite);

            Property startrow = new Property
            {
                Name = "startrow",
                MetaData = new MetaData("Start Row", "Start Row"),
                SoType = SourceCode.SmartObjects.Services.ServiceSDK.Types.SoType.Number,
                Type = "System.Int32"
            };
            ContainerProperties.Add(startrow);

            Property rowlimit = new Property
            {
                Name = "rowlimit",
                MetaData = new MetaData("Row Limit", "Row Limit"),
                SoType = SourceCode.SmartObjects.Services.ServiceSDK.Types.SoType.Number,
                Type = "System.Int32"
            };
            ContainerProperties.Add(rowlimit);

            Property sort = new Property
            {
                Name = "sort",
                MetaData = new MetaData("Sort", "Sort"),
                SoType = SourceCode.SmartObjects.Services.ServiceSDK.Types.SoType.Text,
                Type = "System.String"
            };
            ContainerProperties.Add(sort);

            Property enablephonetic = new Property
            {
                Name = "enablephonetic",
                MetaData = new MetaData("Enable Phonetic", "Enable Phonetic"),
                SoType = SourceCode.SmartObjects.Services.ServiceSDK.Types.SoType.YesNo,
                Type = "System.Boolean"
            };
            ContainerProperties.Add(enablephonetic);

            Property enablenicknames = new Property
            {
                Name = "enablenicknames",
                MetaData = new MetaData("Enable Nicknames", "Enable Nicknames"),
                SoType = SourceCode.SmartObjects.Services.ServiceSDK.Types.SoType.YesNo,
                Type = "System.Boolean"
            };
            ContainerProperties.Add(enablenicknames);


            Property enablestemming = new Property
            {
                Name = "enablestemming",
                MetaData = new MetaData("Enable Stemming", "Enable Stemming"),
                SoType = SourceCode.SmartObjects.Services.ServiceSDK.Types.SoType.YesNo,
                Type = "System.Boolean"
            };
            ContainerProperties.Add(enablestemming);

            Property trimduplicates = new Property
            {
                Name = "trimduplicates",
                MetaData = new MetaData("Trim Duplicates", "Trip Duplicates"),
                SoType = SourceCode.SmartObjects.Services.ServiceSDK.Types.SoType.YesNo,
                Type = "System.Boolean"
            };
            ContainerProperties.Add(trimduplicates);

            Property enablequeryrules = new Property
            {
                Name = "enablequeryrules",
                MetaData = new MetaData("Enable Query Rules", "Enable Query Rules"),
                SoType = SourceCode.SmartObjects.Services.ServiceSDK.Types.SoType.YesNo,
                Type = "System.Boolean"
            };
            ContainerProperties.Add(enablequeryrules);

            Property processbestbets = new Property
            {
                Name = "processbestbets",
                MetaData = new MetaData("Process Best Bets", "Process Best Bets"),
                SoType = SourceCode.SmartObjects.Services.ServiceSDK.Types.SoType.YesNo,
                Type = "System.Boolean"
            };
            ContainerProperties.Add(processbestbets);

            Property processpersonal = new Property
            {
                Name = "processpersonal",
                MetaData = new MetaData("Process Personal", "Process Personal"),
                SoType = SourceCode.SmartObjects.Services.ServiceSDK.Types.SoType.YesNo,
                Type = "System.Boolean"
            };
            ContainerProperties.Add(processpersonal);

            Property sourceid = new Property
            {
                Name = "sourceid",
                MetaData = new MetaData("Source Id", "Source Id"),
                SoType = SourceCode.SmartObjects.Services.ServiceSDK.Types.SoType.Guid,
                Type = "System.Guid"
            };
            ContainerProperties.Add(sourceid);

            Property sourcename = new Property
            {
                Name = "sourcename",
                MetaData = new MetaData("Source Name", "Source Name"),
                SoType = SourceCode.SmartObjects.Services.ServiceSDK.Types.SoType.Text,
                Type = "System.String"
            };
            ContainerProperties.Add(sourcename);


            Property fileextensions = new Property
            {
                Name = "fileextensions",
                MetaData = new MetaData("File Extensions", "File Extensions"),
                SoType = SourceCode.SmartObjects.Services.ServiceSDK.Types.SoType.Text,
                Type = "System.String"
            };
            ContainerProperties.Add(fileextensions);

            return ContainerProperties;
        }

        public static List<Property> GetSearchResultSummaryProperties()
        {
            List<Property> ContainerProperties = new List<Property>();

            Property executiontime = new Property
            {
                Name = "executiontime",
                MetaData = new MetaData("Execution Time", "Execution Time"),
                SoType = SourceCode.SmartObjects.Services.ServiceSDK.Types.SoType.Number,
                Type = "System.Int32"
            };
            ContainerProperties.Add(executiontime);

            Property resultrows = new Property
            {
                Name = "resultrows",
                MetaData = new MetaData("Result Rows", "Result Rows"),
                SoType = SourceCode.SmartObjects.Services.ServiceSDK.Types.SoType.Number,
                Type = "System.Int32"
            };
            ContainerProperties.Add(resultrows);

            Property totalrows = new Property
            {
                Name = "totalrows",
                MetaData = new MetaData("Total Rows", "Total Rows"),
                SoType = SourceCode.SmartObjects.Services.ServiceSDK.Types.SoType.Number,
                Type = "System.Int32"
            };
            ContainerProperties.Add(totalrows);

            //Property ResultsTitle = new Property
            //{
            //    Name = "resulttitle",
            //    MetaData = new MetaData("Result Title", "Results Title"),
            //    SoType = SourceCode.SmartObjects.Services.ServiceSDK.Types.SoType.Text,
            //    Type = "System.String"
            //};
            //ContainerProperties.Add(ResultsTitle);

            //Property ResultsTitleUrl = new Property
            //{
            //    Name = "resulttitleurl",
            //    MetaData = new MetaData("Result Title Url", "Result Title Url"),
            //    SoType = SourceCode.SmartObjects.Services.ServiceSDK.Types.SoType.Text,
            //    Type = "System.String"
            //};
            //ContainerProperties.Add(ResultsTitleUrl);

            //Property SpellingSuggestions = new Property
            //{
            //    Name = "spellingsuggestions",
            //    MetaData = new MetaData("Spelling Suggestions", "Spelling Suggestions"),
            //    SoType = SourceCode.SmartObjects.Services.ServiceSDK.Types.SoType.Text,
            //    Type = "System.String"
            //};
            //ContainerProperties.Add(SpellingSuggestions);

            //Property TableType = new Property
            //{
            //    Name = "tabletype",
            //    MetaData = new MetaData("Table Type", "Table Type"),
            //    SoType = SourceCode.SmartObjects.Services.ServiceSDK.Types.SoType.Text,
            //    Type = "System.String"
            //};
            //ContainerProperties.Add(TableType);


            Property SerializedResults = new Property
            {
                Name = "serializedresults",
                MetaData = new MetaData("Serialized Results", "Serialized Results"),
                SoType = SourceCode.SmartObjects.Services.ServiceSDK.Types.SoType.Memo,
                Type = "System.String"
            };
            ContainerProperties.Add(SerializedResults);

            return ContainerProperties;
        }

        public static List<Property> GetSearchResultsProperties()
        {
            List<Property> ContainerProperties = new List<Property>();

            ContainerProperties.AddRange(GetStandardSearchReturnProperties());
            ContainerProperties.AddRange(GetSearchResultReturnProperties());

            return ContainerProperties.OrderBy(p => p.MetaData.DisplayName).ToList();
        }

        public static List<Property> GetStandardSearchReturnProperties()
        {
            List<Property> ContainerProperties = new List<Property>();

            Property rank = new Property
            {
                Name = "rank",
                MetaData = new MetaData("Rank", "Rank"),
                SoType = SourceCode.SmartObjects.Services.ServiceSDK.Types.SoType.Text,
                Type = "System.String"
            };
            ContainerProperties.Add(rank);

            Property lastmodifiedtime = new Property
            {
                Name = "lastmodifiedtime",
                MetaData = new MetaData("LastModifiedTime", "LastModifiedTime"),
                SoType = SourceCode.SmartObjects.Services.ServiceSDK.Types.SoType.DateTime,
                Type = "System.DateTime"
            };
            ContainerProperties.Add(lastmodifiedtime);

            Property path = new Property
            {
                Name = "path",
                MetaData = new MetaData("Path", "Path"),
                SoType = SourceCode.SmartObjects.Services.ServiceSDK.Types.SoType.Text,
                Type = "System.String"
            };
            ContainerProperties.Add(path);

            Property docid = new Property
            {
                Name = "docid",
                MetaData = new MetaData("DocId", "DocId"),
                SoType = SourceCode.SmartObjects.Services.ServiceSDK.Types.SoType.Text,
                Type = "System.String"
            };
            ContainerProperties.Add(docid);

            Property workid = new Property
            {
                Name = "workid",
                MetaData = new MetaData("WorkId", "WorkId"),
                SoType = SourceCode.SmartObjects.Services.ServiceSDK.Types.SoType.Text,
                Type = "System.String"
            };
            ContainerProperties.Add(workid);

            Property originalpath = new Property
            {
                Name = "originalpath",
                MetaData = new MetaData("OriginalPath", "OriginalPath"),
                SoType = SourceCode.SmartObjects.Services.ServiceSDK.Types.SoType.Text,
                Type = "System.String"
            };
            ContainerProperties.Add(originalpath);

            return ContainerProperties;
        }

        public static List<Property> GetSearchResultReturnProperties()
        {
            List<Property> ContainerProperties = new List<Property>();            

            Property title = new Property
            {
                Name = "title",
                MetaData = new MetaData("Title", "Title"),
                SoType = SourceCode.SmartObjects.Services.ServiceSDK.Types.SoType.Text,
                Type = "System.String"
            };
            ContainerProperties.Add(title);

            Property author = new Property
            {
                Name = "author",
                MetaData = new MetaData("Author", "Author"),
                SoType = SourceCode.SmartObjects.Services.ServiceSDK.Types.SoType.Text,
                Type = "System.String"
            };
            ContainerProperties.Add(author);

            Property size = new Property
            {
                Name = "size",
                MetaData = new MetaData("Size", "Size"),
                SoType = SourceCode.SmartObjects.Services.ServiceSDK.Types.SoType.Number,
                Type = "System.Int32"
            };
            ContainerProperties.Add(size);

            Property description = new Property
            {
                Name = "description",
                MetaData = new MetaData("Description", "Description"),
                SoType = SourceCode.SmartObjects.Services.ServiceSDK.Types.SoType.Memo,
                Type = "System.String"
            };
            ContainerProperties.Add(description);

            Property write = new Property
            {
                Name = "write",
                MetaData = new MetaData("Write", "Write"),
                SoType = SourceCode.SmartObjects.Services.ServiceSDK.Types.SoType.DateTime,
                Type = "System.DateTime"
            };
            ContainerProperties.Add(write);

            //Property collapsingstatus = new Property
            //{
            //    Name = "collapsingstatus",
            //    MetaData = new MetaData("CollapsingStatus", "CollapsingStatus"),
            //    SoType = SourceCode.SmartObjects.Services.ServiceSDK.Types.SoType.Text,
            //    Type = "System.String"
            //};
            //ContainerProperties.Add(collapsingstatus);

            Property hithighlightedsummary = new Property
            {
                Name = "hithighlightedsummary",
                MetaData = new MetaData("HitHighlightedSummary", "HitHighlightedSummary"),
                SoType = SourceCode.SmartObjects.Services.ServiceSDK.Types.SoType.Text,
                Type = "System.String"
            };
            ContainerProperties.Add(hithighlightedsummary);

            Property hithighlightedproperties = new Property
            {
                Name = "hithighlightedproperties",
                MetaData = new MetaData("HitHighlightedProperties", "HitHighlightedProperties"),
                SoType = SourceCode.SmartObjects.Services.ServiceSDK.Types.SoType.Text,
                Type = "System.String"
            };
            ContainerProperties.Add(hithighlightedproperties);

            //Property contentclass = new Property
            //{
            //    Name = "contentclass",
            //    MetaData = new MetaData("contentclass", "contentclass"),
            //    SoType = SourceCode.SmartObjects.Services.ServiceSDK.Types.SoType.Text,
            //    Type = "System.String"
            //};
            //ContainerProperties.Add(contentclass);

            Property picturethumbnailurl = new Property
            {
                Name = "picturethumbnailurl",
                MetaData = new MetaData("PictureThumbnailURL", "PictureThumbnailURL"),
                SoType = SourceCode.SmartObjects.Services.ServiceSDK.Types.SoType.Text,
                Type = "System.String"
            };
            ContainerProperties.Add(picturethumbnailurl);

            Property serverredirectedurl = new Property
            {
                Name = "serverredirectedurl",
                MetaData = new MetaData("ServerRedirectedURL", "ServerRedirectedURL"),
                SoType = SourceCode.SmartObjects.Services.ServiceSDK.Types.SoType.Text,
                Type = "System.String"
            };
            ContainerProperties.Add(serverredirectedurl);

            Property serverredirectedembedurl = new Property
            {
                Name = "serverredirectedembedurl",
                MetaData = new MetaData("ServerRedirectedEmbedURL", "ServerRedirectedEmbedURL"),
                SoType = SourceCode.SmartObjects.Services.ServiceSDK.Types.SoType.Text,
                Type = "System.String"
            };
            ContainerProperties.Add(serverredirectedembedurl);

            Property serverredirectedpreviewurl = new Property
            {
                Name = "serverredirectedpreviewurl",
                MetaData = new MetaData("ServerRedirectedPreviewURL", "ServerRedirectedPreviewURL"),
                SoType = SourceCode.SmartObjects.Services.ServiceSDK.Types.SoType.Text,
                Type = "System.String"
            };
            ContainerProperties.Add(serverredirectedpreviewurl);

            Property fileextension = new Property
            {
                Name = "fileextension",
                MetaData = new MetaData("FileExtension", "FileExtension"),
                SoType = SourceCode.SmartObjects.Services.ServiceSDK.Types.SoType.Text,
                Type = "System.String"
            };
            ContainerProperties.Add(fileextension);

            Property contenttypeid = new Property
            {
                Name = "contenttypeid",
                MetaData = new MetaData("ContentTypeId", "ContentTypeId"),
                SoType = SourceCode.SmartObjects.Services.ServiceSDK.Types.SoType.Text,
                Type = "System.String"
            };
            ContainerProperties.Add(contenttypeid);

            Property parentlink = new Property
            {
                Name = "parentlink",
                MetaData = new MetaData("ParentLink", "ParentLink"),
                SoType = SourceCode.SmartObjects.Services.ServiceSDK.Types.SoType.Text,
                Type = "System.String"
            };
            ContainerProperties.Add(parentlink);

            Property viewslifetime = new Property
            {
                Name = "viewslifetime",
                MetaData = new MetaData("ViewsLifeTime", "ViewsLifeTime"),
                SoType = SourceCode.SmartObjects.Services.ServiceSDK.Types.SoType.Text,
                Type = "System.String"
            };
            ContainerProperties.Add(viewslifetime);

            Property viewsrecent = new Property
            {
                Name = "viewsrecent",
                MetaData = new MetaData("ViewsRecent", "ViewsRecent"),
                SoType = SourceCode.SmartObjects.Services.ServiceSDK.Types.SoType.Text,
                Type = "System.String"
            };
            ContainerProperties.Add(viewsrecent);

            Property sectionnames = new Property
            {
                Name = "sectionnames",
                MetaData = new MetaData("SectionNames", "SectionNames"),
                SoType = SourceCode.SmartObjects.Services.ServiceSDK.Types.SoType.Text,
                Type = "System.String"
            };
            ContainerProperties.Add(sectionnames);

            //Property sectionindexes = new Property
            //{
            //    Name = "sectionindexes",
            //    MetaData = new MetaData("SectionIndexes", "SectionIndexes"),
            //    SoType = SourceCode.SmartObjects.Services.ServiceSDK.Types.SoType.Text,
            //    Type = "System.String"
            //};
            //ContainerProperties.Add(sectionindexes);

            Property sitelogo = new Property
            {
                Name = "sitelogo",
                MetaData = new MetaData("SiteLogo", "SiteLogo"),
                SoType = SourceCode.SmartObjects.Services.ServiceSDK.Types.SoType.Text,
                Type = "System.String"
            };
            ContainerProperties.Add(sitelogo);

            Property sitedescription = new Property
            {
                Name = "sitedescription",
                MetaData = new MetaData("SiteDescription", "SiteDescription"),
                SoType = SourceCode.SmartObjects.Services.ServiceSDK.Types.SoType.Memo,
                Type = "System.String"
            };
            ContainerProperties.Add(sitedescription);

            Property deeplinks = new Property
            {
                Name = "deeplinks",
                MetaData = new MetaData("deeplinks", "deeplinks"),
                SoType = SourceCode.SmartObjects.Services.ServiceSDK.Types.SoType.Text,
                Type = "System.String"
            };
            ContainerProperties.Add(deeplinks);

            Property importance = new Property
            {
                Name = "importance",
                MetaData = new MetaData("importance", "importance"),
                SoType = SourceCode.SmartObjects.Services.ServiceSDK.Types.SoType.Text,
                Type = "System.String"
            };
            ContainerProperties.Add(importance);

            Property sitename = new Property
            {
                Name = "sitename",
                MetaData = new MetaData("SiteName", "SiteName"),
                SoType = SourceCode.SmartObjects.Services.ServiceSDK.Types.SoType.Text,
                Type = "System.String"
            };
            ContainerProperties.Add(sitename);

            Property isdocument = new Property
            {
                Name = "isdocument",
                MetaData = new MetaData("IsDocument", "IsDocument"),
                SoType = SourceCode.SmartObjects.Services.ServiceSDK.Types.SoType.YesNo,
                Type = "System.Boolean"
            };
            ContainerProperties.Add(isdocument);

            Property filetype = new Property
            {
                Name = "filetype",
                MetaData = new MetaData("FileType", "FileType"),
                SoType = SourceCode.SmartObjects.Services.ServiceSDK.Types.SoType.Text,
                Type = "System.String"
            };
            ContainerProperties.Add(filetype);

            Property iscontainer = new Property
            {
                Name = "iscontainer",
                MetaData = new MetaData("IsContainer", "IsContainer"),
                SoType = SourceCode.SmartObjects.Services.ServiceSDK.Types.SoType.YesNo,
                Type = "System.Boolean"
            };
            ContainerProperties.Add(iscontainer);

            Property webtemplate = new Property
            {
                Name = "webtemplate",
                MetaData = new MetaData("WebTemplate", "WebTemplate"),
                SoType = SourceCode.SmartObjects.Services.ServiceSDK.Types.SoType.Text,
                Type = "System.String"
            };
            ContainerProperties.Add(webtemplate);

            Property secondaryfileextension = new Property
            {
                Name = "secondaryfileextension",
                MetaData = new MetaData("SecondaryFileExtension", "SecondaryFileExtension"),
                SoType = SourceCode.SmartObjects.Services.ServiceSDK.Types.SoType.Text,
                Type = "System.String"
            };
            ContainerProperties.Add(secondaryfileextension);

            //Property docaclmeta = new Property
            //{
            //    Name = "docaclmeta",
            //    MetaData = new MetaData("docaclmeta", "docaclmeta"),
            //    SoType = SourceCode.SmartObjects.Services.ServiceSDK.Types.SoType.Text,
            //    Type = "System.String"
            //};
            //ContainerProperties.Add(docaclmeta);

            //Property editorowsuser = new Property
            //{
            //    Name = "editorowsuser",
            //    MetaData = new MetaData("EditorOWSUSER", "EditorOWSUSER"),
            //    SoType = SourceCode.SmartObjects.Services.ServiceSDK.Types.SoType.Text,
            //    Type = "System.String"
            //};
            //ContainerProperties.Add(editorowsuser);

            Property displayauthor = new Property
            {
                Name = "displayauthor",
                MetaData = new MetaData("DisplayAuthor", "DisplayAuthor"),
                SoType = SourceCode.SmartObjects.Services.ServiceSDK.Types.SoType.Text,
                Type = "System.String"
            };
            ContainerProperties.Add(displayauthor);

            Property replycount = new Property
            {
                Name = "replycount",
                MetaData = new MetaData("ReplyCount", "ReplyCount"),
                SoType = SourceCode.SmartObjects.Services.ServiceSDK.Types.SoType.Text,
                Type = "System.String"
            };
            ContainerProperties.Add(replycount);

            Property likescount = new Property
            {
                Name = "likescount",
                MetaData = new MetaData("LikesCount", "LikesCount"),
                SoType = SourceCode.SmartObjects.Services.ServiceSDK.Types.SoType.Text,
                Type = "System.String"
            };
            ContainerProperties.Add(likescount);

            Property created = new Property
            {
                Name = "created",
                MetaData = new MetaData("Created", "Created"),
                SoType = SourceCode.SmartObjects.Services.ServiceSDK.Types.SoType.DateTime,
                Type = "System.DateTime"
            };
            ContainerProperties.Add(created);

            Property listitemid = new Property
            {
                Name = "listitemid",
                MetaData = new MetaData("ListItemID", "ListItemID"),
                SoType = SourceCode.SmartObjects.Services.ServiceSDK.Types.SoType.Text,
                Type = "System.String"
            };
            ContainerProperties.Add(listitemid);

            Property fullpostbody = new Property
            {
                Name = "fullpostbody",
                MetaData = new MetaData("FullPostBody", "FullPostBody"),
                SoType = SourceCode.SmartObjects.Services.ServiceSDK.Types.SoType.Memo,
                Type = "System.String"
            };
            ContainerProperties.Add(fullpostbody);

            Property postauthor = new Property
            {
                Name = "postauthor",
                MetaData = new MetaData("PostAuthor", "PostAuthor"),
                SoType = SourceCode.SmartObjects.Services.ServiceSDK.Types.SoType.Text,
                Type = "System.String"
            };
            ContainerProperties.Add(postauthor);

            //Property rootpostownerid = new Property
            //{
            //    Name = "rootpostownerid",
            //    MetaData = new MetaData("RootPostOwnerID", "RootPostOwnerID"),
            //    SoType = SourceCode.SmartObjects.Services.ServiceSDK.Types.SoType.Text,
            //    Type = "System.String"
            //};
            //ContainerProperties.Add(rootpostownerid);

            //Property rootpostid = new Property
            //{
            //    Name = "rootpostid",
            //    MetaData = new MetaData("RootPostID", "RootPostID"),
            //    SoType = SourceCode.SmartObjects.Services.ServiceSDK.Types.SoType.Text,
            //    Type = "System.String"
            //};
            //ContainerProperties.Add(rootpostid);

            Property attachmenttype = new Property
            {
                Name = "attachmenttype",
                MetaData = new MetaData("AttachmentType", "AttachmentType"),
                SoType = SourceCode.SmartObjects.Services.ServiceSDK.Types.SoType.Text,
                Type = "System.String"
            };
            ContainerProperties.Add(attachmenttype);

            Property attachmenturi = new Property
            {
                Name = "attachmenturi",
                MetaData = new MetaData("AttachmentURI", "AttachmentURI"),
                SoType = SourceCode.SmartObjects.Services.ServiceSDK.Types.SoType.Text,
                Type = "System.String"
            };
            ContainerProperties.Add(attachmenturi);

            Property microblogtype = new Property
            {
                Name = "microblogtype",
                MetaData = new MetaData("MicroBlogType", "MicroBlogType"),
                SoType = SourceCode.SmartObjects.Services.ServiceSDK.Types.SoType.Text,
                Type = "System.String"
            };
            ContainerProperties.Add(microblogtype);

            Property modifiedby = new Property
            {
                Name = "modifiedby",
                MetaData = new MetaData("ModifiedBy", "ModifiedBy"),
                SoType = SourceCode.SmartObjects.Services.ServiceSDK.Types.SoType.Text,
                Type = "System.String"
            };
            ContainerProperties.Add(modifiedby);

            //Property rootpostuniqueid = new Property
            //{
            //    Name = "rootpostuniqueid",
            //    MetaData = new MetaData("RootPostUniqueID", "RootPostUniqueID"),
            //    SoType = SourceCode.SmartObjects.Services.ServiceSDK.Types.SoType.Text,
            //    Type = "System.String"
            //};
            //ContainerProperties.Add(rootpostuniqueid);

            Property tags = new Property
            {
                Name = "tags",
                MetaData = new MetaData("Tags", "Tags"),
                SoType = SourceCode.SmartObjects.Services.ServiceSDK.Types.SoType.Text,
                Type = "System.String"
            };
            ContainerProperties.Add(tags);

            //Property resulttypeidlist = new Property
            //{
            //    Name = "resulttypeidlist",
            //    MetaData = new MetaData("ResultTypeIdList", "ResultTypeIdList"),
            //    SoType = SourceCode.SmartObjects.Services.ServiceSDK.Types.SoType.Text,
            //    Type = "System.String"
            //};
            //ContainerProperties.Add(resulttypeidlist);

            //Property partitionid = new Property
            //{
            //    Name = "partitionid",
            //    MetaData = new MetaData("PartitionId", "PartitionId"),
            //    SoType = SourceCode.SmartObjects.Services.ServiceSDK.Types.SoType.Text,
            //    Type = "System.String"
            //};
            //ContainerProperties.Add(partitionid);

            //Property urlzone = new Property
            //{
            //    Name = "urlzone",
            //    MetaData = new MetaData("UrlZone", "UrlZone"),
            //    SoType = SourceCode.SmartObjects.Services.ServiceSDK.Types.SoType.Text,
            //    Type = "System.String"
            //};
            //ContainerProperties.Add(urlzone);

            //Property aamenabledmanagedproperties = new Property
            //{
            //    Name = "aamenabledmanagedproperties",
            //    MetaData = new MetaData("AAMEnabledManagedProperties", "AAMEnabledManagedProperties"),
            //    SoType = SourceCode.SmartObjects.Services.ServiceSDK.Types.SoType.Text,
            //    Type = "System.String"
            //};
            //ContainerProperties.Add(aamenabledmanagedproperties);

            //Property resulttypeid = new Property
            //{
            //    Name = "resulttypeid",
            //    MetaData = new MetaData("ResultTypeId", "ResultTypeId"),
            //    SoType = SourceCode.SmartObjects.Services.ServiceSDK.Types.SoType.Text,
            //    Type = "System.String"
            //};
            //ContainerProperties.Add(resulttypeid);

            //Property rendertemplateid = new Property
            //{
            //    Name = "rendertemplateid",
            //    MetaData = new MetaData("RenderTemplateId", "RenderTemplateId"),
            //    SoType = SourceCode.SmartObjects.Services.ServiceSDK.Types.SoType.Text,
            //    Type = "System.String"
            //};
            //ContainerProperties.Add(rendertemplateid);

            //Property pisearchresultid = new Property
            //{
            //    Name = "pisearchresultid",
            //    MetaData = new MetaData("piSearchResultId", "piSearchResultId"),
            //    SoType = SourceCode.SmartObjects.Services.ServiceSDK.Types.SoType.Text,
            //    Type = "System.String"
            //};
            //ContainerProperties.Add(pisearchresultid);

            return ContainerProperties.OrderBy(p => p.MetaData.DisplayName).ToList();

        }

        public static List<Property> GetUserSearchResultProperties()
        {

            List<Property> ContainerProperties = new List<Property>();

            Property rank = new Property
            {
                Name = "rank",
                MetaData = new MetaData("Rank", "Rank"),
                SoType = SourceCode.SmartObjects.Services.ServiceSDK.Types.SoType.Text,
                Type = "System.String"
            };
            ContainerProperties.Add(rank);

            Property docid = new Property
            {
                Name = "docid",
                MetaData = new MetaData("DocId", "DocId"),
                SoType = SourceCode.SmartObjects.Services.ServiceSDK.Types.SoType.Text,
                Type = "System.String"
            };
            ContainerProperties.Add(docid);

            Property aboutme = new Property
            {
                Name = "aboutme",
                MetaData = new MetaData("AboutMe", "AboutMe"),
                SoType = SourceCode.SmartObjects.Services.ServiceSDK.Types.SoType.Memo,
                Type = "System.String"
            };
            ContainerProperties.Add(aboutme);

            Property accountname = new Property
            {
                Name = "accountname",
                MetaData = new MetaData("AccountName", "AccountName"),
                SoType = SourceCode.SmartObjects.Services.ServiceSDK.Types.SoType.Text,
                Type = "System.String"
            };
            ContainerProperties.Add(accountname);

            Property baseofficelocation = new Property
            {
                Name = "baseofficelocation",
                MetaData = new MetaData("BaseOfficeLocation", "BaseOfficeLocation"),
                SoType = SourceCode.SmartObjects.Services.ServiceSDK.Types.SoType.Text,
                Type = "System.String"
            };
            ContainerProperties.Add(baseofficelocation);

            Property department = new Property
            {
                Name = "department",
                MetaData = new MetaData("Department", "Department"),
                SoType = SourceCode.SmartObjects.Services.ServiceSDK.Types.SoType.Text,
                Type = "System.String"
            };
            ContainerProperties.Add(department);

            Property hithighlightedproperties = new Property
            {
                Name = "hithighlightedproperties",
                MetaData = new MetaData("HitHighlightedProperties", "HitHighlightedProperties"),
                SoType = SourceCode.SmartObjects.Services.ServiceSDK.Types.SoType.Memo,
                Type = "System.String"
            };
            ContainerProperties.Add(hithighlightedproperties);

            Property interests = new Property
            {
                Name = "interests",
                MetaData = new MetaData("Interests", "Interests"),
                SoType = SourceCode.SmartObjects.Services.ServiceSDK.Types.SoType.Memo,
                Type = "System.String"
            };
            ContainerProperties.Add(interests);

            Property jobtitle = new Property
            {
                Name = "jobtitle",
                MetaData = new MetaData("JobTitle", "JobTitle"),
                SoType = SourceCode.SmartObjects.Services.ServiceSDK.Types.SoType.Text,
                Type = "System.String"
            };
            ContainerProperties.Add(jobtitle);

            Property lastmodifiedtime = new Property
            {
                Name = "lastmodifiedtime",
                MetaData = new MetaData("LastModifiedTime", "LastModifiedTime"),
                SoType = SourceCode.SmartObjects.Services.ServiceSDK.Types.SoType.DateTime,
                Type = "System.DateTime"
            };
            ContainerProperties.Add(lastmodifiedtime);

            Property memberships = new Property
            {
                Name = "memberships",
                MetaData = new MetaData("Memberships", "Memberships"),
                SoType = SourceCode.SmartObjects.Services.ServiceSDK.Types.SoType.Text,
                Type = "System.String"
            };
            ContainerProperties.Add(memberships);

            Property pastprojects = new Property
            {
                Name = "pastprojects",
                MetaData = new MetaData("PastProjects", "PastProjects"),
                SoType = SourceCode.SmartObjects.Services.ServiceSDK.Types.SoType.Memo,
                Type = "System.String"
            };
            ContainerProperties.Add(pastprojects);

            //Property path = new Property
            //{
            //    Name = "path",
            //    MetaData = new MetaData("Path", "Path"),
            //    SoType = SourceCode.SmartObjects.Services.ServiceSDK.Types.SoType.Text,
            //    Type = "System.String"
            //};
            //ContainerProperties.Add(path);

            Property pictureurl = new Property
            {
                Name = "pictureurl",
                MetaData = new MetaData("PictureURL", "PictureURL"),
                SoType = SourceCode.SmartObjects.Services.ServiceSDK.Types.SoType.Text,
                Type = "System.String"
            };
            ContainerProperties.Add(pictureurl);

            Property preferredname = new Property
            {
                Name = "preferredname",
                MetaData = new MetaData("PreferredName", "PreferredName"),
                SoType = SourceCode.SmartObjects.Services.ServiceSDK.Types.SoType.Text,
                Type = "System.String"
            };
            ContainerProperties.Add(preferredname);

            Property responsibilities = new Property
            {
                Name = "responsibilities",
                MetaData = new MetaData("Responsibilities", "Responsibilities"),
                SoType = SourceCode.SmartObjects.Services.ServiceSDK.Types.SoType.Memo,
                Type = "System.String"
            };
            ContainerProperties.Add(responsibilities);

            Property schools = new Property
            {
                Name = "schools",
                MetaData = new MetaData("Schools", "Schools"),
                SoType = SourceCode.SmartObjects.Services.ServiceSDK.Types.SoType.Memo,
                Type = "System.String"
            };
            ContainerProperties.Add(schools);

            //Property serviceapplicationid = new Property
            //{
            //    Name = "serviceapplicationid",
            //    MetaData = new MetaData("ServiceApplicationID", "ServiceApplicationID"),
            //    SoType = SourceCode.SmartObjects.Services.ServiceSDK.Types.SoType.Text,
            //    Type = "System.String"
            //};
            //ContainerProperties.Add(serviceapplicationid);

            Property sipaddress = new Property
            {
                Name = "sipaddress",
                MetaData = new MetaData("SipAddress", "SipAddress"),
                SoType = SourceCode.SmartObjects.Services.ServiceSDK.Types.SoType.Text,
                Type = "System.String"
            };
            ContainerProperties.Add(sipaddress);

            Property skills = new Property
            {
                Name = "skills",
                MetaData = new MetaData("Skills", "Skills"),
                SoType = SourceCode.SmartObjects.Services.ServiceSDK.Types.SoType.Memo,
                Type = "System.String"
            };
            ContainerProperties.Add(skills);

            Property userprofile_guid = new Property
            {
                Name = "userprofile_guid",
                MetaData = new MetaData("UserProfile_GUID", "UserProfile_GUID"),
                SoType = SourceCode.SmartObjects.Services.ServiceSDK.Types.SoType.Guid,
                Type = "System.Guid"
            };
            ContainerProperties.Add(userprofile_guid);

            Property workemail = new Property
            {
                Name = "workemail",
                MetaData = new MetaData("WorkEmail", "WorkEmail"),
                SoType = SourceCode.SmartObjects.Services.ServiceSDK.Types.SoType.Text,
                Type = "System.String"
            };
            ContainerProperties.Add(workemail);

            Property workid = new Property
            {
                Name = "workid",
                MetaData = new MetaData("WorkId", "WorkId"),
                SoType = SourceCode.SmartObjects.Services.ServiceSDK.Types.SoType.Text,
                Type = "System.String"
            };
            ContainerProperties.Add(workid);

            Property yomidisplayname = new Property
            {
                Name = "yomidisplayname",
                MetaData = new MetaData("YomiDisplayName", "YomiDisplayName"),
                SoType = SourceCode.SmartObjects.Services.ServiceSDK.Types.SoType.Text,
                Type = "System.String"
            };
            ContainerProperties.Add(yomidisplayname);

            //Property docaclmeta = new Property
            //{
            //    Name = "docaclmeta",
            //    MetaData = new MetaData("docaclmeta", "docaclmeta"),
            //    SoType = SourceCode.SmartObjects.Services.ServiceSDK.Types.SoType.Text,
            //    Type = "System.String"
            //};
            //ContainerProperties.Add(docaclmeta);

            //Property originalpath = new Property
            //{
            //    Name = "originalpath",
            //    MetaData = new MetaData("OriginalPath", "OriginalPath"),
            //    SoType = SourceCode.SmartObjects.Services.ServiceSDK.Types.SoType.Text,
            //    Type = "System.String"
            //};
            //ContainerProperties.Add(originalpath);

            Property resulttypeidlist = new Property
            {
                Name = "resulttypeidlist",
                MetaData = new MetaData("ResultTypeIdList", "ResultTypeIdList"),
                SoType = SourceCode.SmartObjects.Services.ServiceSDK.Types.SoType.Text,
                Type = "System.String"
            };
            ContainerProperties.Add(resulttypeidlist);

            //Property partitionid = new Property
            //{
            //    Name = "partitionid",
            //    MetaData = new MetaData("PartitionId", "PartitionId"),
            //    SoType = SourceCode.SmartObjects.Services.ServiceSDK.Types.SoType.Text,
            //    Type = "System.String"
            //};
            //ContainerProperties.Add(partitionid);

            Property urlzone = new Property
            {
                Name = "urlzone",
                MetaData = new MetaData("UrlZone", "UrlZone"),
                SoType = SourceCode.SmartObjects.Services.ServiceSDK.Types.SoType.Text,
                Type = "System.String"
            };
            ContainerProperties.Add(urlzone);

            //Property aamenabledmanagedproperties = new Property
            //{
            //    Name = "aamenabledmanagedproperties",
            //    MetaData = new MetaData("AAMEnabledManagedProperties", "AAMEnabledManagedProperties"),
            //    SoType = SourceCode.SmartObjects.Services.ServiceSDK.Types.SoType.Text,
            //    Type = "System.String"
            //};
            //ContainerProperties.Add(aamenabledmanagedproperties);

            //Property resulttypeid = new Property
            //{
            //    Name = "resulttypeid",
            //    MetaData = new MetaData("ResultTypeId", "ResultTypeId"),
            //    SoType = SourceCode.SmartObjects.Services.ServiceSDK.Types.SoType.Text,
            //    Type = "System.String"
            //};
            //ContainerProperties.Add(resulttypeid);

            Property editprofileurl = new Property
            {
                Name = "editprofileurl",
                MetaData = new MetaData("EditProfileUrl", "EditProfileUrl"),
                SoType = SourceCode.SmartObjects.Services.ServiceSDK.Types.SoType.Text,
                Type = "System.String"
            };
            ContainerProperties.Add(editprofileurl);

            Property profileviewslastmonth = new Property
            {
                Name = "profileviewslastmonth",
                MetaData = new MetaData("ProfileViewsLastMonth", "ProfileViewsLastMonth"),
                SoType = SourceCode.SmartObjects.Services.ServiceSDK.Types.SoType.Text,
                Type = "System.String"
            };
            ContainerProperties.Add(profileviewslastmonth);

            Property profileviewslastweek = new Property
            {
                Name = "profileviewslastweek",
                MetaData = new MetaData("ProfileViewsLastWeek", "ProfileViewsLastWeek"),
                SoType = SourceCode.SmartObjects.Services.ServiceSDK.Types.SoType.Text,
                Type = "System.String"
            };
            ContainerProperties.Add(profileviewslastweek);

            Property profilequeriesfoundyou = new Property
            {
                Name = "profilequeriesfoundyou",
                MetaData = new MetaData("ProfileQueriesFoundYou", "ProfileQueriesFoundYou"),
                SoType = SourceCode.SmartObjects.Services.ServiceSDK.Types.SoType.Text,
                Type = "System.String"
            };
            ContainerProperties.Add(profilequeriesfoundyou);

            //Property rendertemplateid = new Property
            //{
            //    Name = "rendertemplateid",
            //    MetaData = new MetaData("RenderTemplateId", "RenderTemplateId"),
            //    SoType = SourceCode.SmartObjects.Services.ServiceSDK.Types.SoType.Text,
            //    Type = "System.String"
            //};
            //ContainerProperties.Add(rendertemplateid);

            //Property pisearchresultid = new Property
            //{
            //    Name = "pisearchresultid",
            //    MetaData = new MetaData("piSearchResultId", "piSearchResultId"),
            //    SoType = SourceCode.SmartObjects.Services.ServiceSDK.Types.SoType.Text,
            //    Type = "System.String"
            //};
            //ContainerProperties.Add(pisearchresultid);


            return ContainerProperties.OrderBy(p => p.MetaData.DisplayName).ToList();
        }

        public static List<Property> GetGraphSearchResultProperties()
        {
            List<Property> ContainerProperties = new List<Property>();


            Property edges = new Property
            {
                Name = "graphedges",
                MetaData = new MetaData("Edges", "Edges"),
                SoType = SourceCode.SmartObjects.Services.ServiceSDK.Types.SoType.Memo,
                Type = "System.String"
            };
            ContainerProperties.Add(edges);


            return ContainerProperties.OrderBy(p => p.MetaData.DisplayName).ToList();
        }
    }
}
