using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using SourceCode.SmartObjects.Services.ServiceSDK;
using SourceCode.SmartObjects.Services.ServiceSDK.Objects;
using SourceCode.SmartObjects.Services.ServiceSDK.Types;
using SourceCode.SharePoint15.Client;
using Microsoft.SharePoint.Client.Search.Query;
using Microsoft.SharePoint.Client;
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
            SPSearchServiceObject.Methods.Add(CreateListSourceIds(SPSearchProps));
            SPSearchServiceObject.Methods.Add(CreateListOtherSourceIds(SPSearchProps));

            serviceBroker.Service.ServiceObjects.Add(SPSearchServiceObject);
        }

        private List<Property> GetSPSearchFullReturnsProperties()
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

            Property workid = new Property
            {
                Name = "workid",
                MetaData = new MetaData("WorkId", "WorkId"),
                SoType = SourceCode.SmartObjects.Services.ServiceSDK.Types.SoType.Text,
                Type = "System.String"
            };
            ContainerProperties.Add(workid);

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

            Property path = new Property
            {
                Name = "path",
                MetaData = new MetaData("Path", "Path"),
                SoType = SourceCode.SmartObjects.Services.ServiceSDK.Types.SoType.Text,
                Type = "System.String"
            };
            ContainerProperties.Add(path);

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

            Property collapsingstatus = new Property
            {
                Name = "collapsingstatus",
                MetaData = new MetaData("CollapsingStatus", "CollapsingStatus"),
                SoType = SourceCode.SmartObjects.Services.ServiceSDK.Types.SoType.Text,
                Type = "System.String"
            };
            ContainerProperties.Add(collapsingstatus);

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

            Property contentclass = new Property
            {
                Name = "contentclass",
                MetaData = new MetaData("contentclass", "contentclass"),
                SoType = SourceCode.SmartObjects.Services.ServiceSDK.Types.SoType.Text,
                Type = "System.String"
            };
            ContainerProperties.Add(contentclass);

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

            Property sectionindexes = new Property
            {
                Name = "sectionindexes",
                MetaData = new MetaData("SectionIndexes", "SectionIndexes"),
                SoType = SourceCode.SmartObjects.Services.ServiceSDK.Types.SoType.Text,
                Type = "System.String"
            };
            ContainerProperties.Add(sectionindexes);

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

            Property lastmodifiedtime = new Property
            {
                Name = "lastmodifiedtime",
                MetaData = new MetaData("LastModifiedTime", "LastModifiedTime"),
                SoType = SourceCode.SmartObjects.Services.ServiceSDK.Types.SoType.DateTime,
                Type = "System.DateTime"
            };
            ContainerProperties.Add(lastmodifiedtime);

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

            Property docaclmeta = new Property
            {
                Name = "docaclmeta",
                MetaData = new MetaData("docaclmeta", "docaclmeta"),
                SoType = SourceCode.SmartObjects.Services.ServiceSDK.Types.SoType.Text,
                Type = "System.String"
            };
            ContainerProperties.Add(docaclmeta);

            Property originalpath = new Property
            {
                Name = "originalpath",
                MetaData = new MetaData("OriginalPath", "OriginalPath"),
                SoType = SourceCode.SmartObjects.Services.ServiceSDK.Types.SoType.Text,
                Type = "System.String"
            };
            ContainerProperties.Add(originalpath);

            Property editorowsuser = new Property
            {
                Name = "editorowsuser",
                MetaData = new MetaData("EditorOWSUSER", "EditorOWSUSER"),
                SoType = SourceCode.SmartObjects.Services.ServiceSDK.Types.SoType.Text,
                Type = "System.String"
            };
            ContainerProperties.Add(editorowsuser);

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

            Property rootpostownerid = new Property
            {
                Name = "rootpostownerid",
                MetaData = new MetaData("RootPostOwnerID", "RootPostOwnerID"),
                SoType = SourceCode.SmartObjects.Services.ServiceSDK.Types.SoType.Text,
                Type = "System.String"
            };
            ContainerProperties.Add(rootpostownerid);

            Property rootpostid = new Property
            {
                Name = "rootpostid",
                MetaData = new MetaData("RootPostID", "RootPostID"),
                SoType = SourceCode.SmartObjects.Services.ServiceSDK.Types.SoType.Text,
                Type = "System.String"
            };
            ContainerProperties.Add(rootpostid);

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

            Property rootpostuniqueid = new Property
            {
                Name = "rootpostuniqueid",
                MetaData = new MetaData("RootPostUniqueID", "RootPostUniqueID"),
                SoType = SourceCode.SmartObjects.Services.ServiceSDK.Types.SoType.Text,
                Type = "System.String"
            };
            ContainerProperties.Add(rootpostuniqueid);

            Property tags = new Property
            {
                Name = "tags",
                MetaData = new MetaData("Tags", "Tags"),
                SoType = SourceCode.SmartObjects.Services.ServiceSDK.Types.SoType.Text,
                Type = "System.String"
            };
            ContainerProperties.Add(tags);

            Property resulttypeidlist = new Property
            {
                Name = "resulttypeidlist",
                MetaData = new MetaData("ResultTypeIdList", "ResultTypeIdList"),
                SoType = SourceCode.SmartObjects.Services.ServiceSDK.Types.SoType.Text,
                Type = "System.String"
            };
            ContainerProperties.Add(resulttypeidlist);

            Property partitionid = new Property
            {
                Name = "partitionid",
                MetaData = new MetaData("PartitionId", "PartitionId"),
                SoType = SourceCode.SmartObjects.Services.ServiceSDK.Types.SoType.Text,
                Type = "System.String"
            };
            ContainerProperties.Add(partitionid);

            Property urlzone = new Property
            {
                Name = "urlzone",
                MetaData = new MetaData("UrlZone", "UrlZone"),
                SoType = SourceCode.SmartObjects.Services.ServiceSDK.Types.SoType.Text,
                Type = "System.String"
            };
            ContainerProperties.Add(urlzone);

            Property aamenabledmanagedproperties = new Property
            {
                Name = "aamenabledmanagedproperties",
                MetaData = new MetaData("AAMEnabledManagedProperties", "AAMEnabledManagedProperties"),
                SoType = SourceCode.SmartObjects.Services.ServiceSDK.Types.SoType.Text,
                Type = "System.String"
            };
            ContainerProperties.Add(aamenabledmanagedproperties);

            Property resulttypeid = new Property
            {
                Name = "resulttypeid",
                MetaData = new MetaData("ResultTypeId", "ResultTypeId"),
                SoType = SourceCode.SmartObjects.Services.ServiceSDK.Types.SoType.Text,
                Type = "System.String"
            };
            ContainerProperties.Add(resulttypeid);

            Property rendertemplateid = new Property
            {
                Name = "rendertemplateid",
                MetaData = new MetaData("RenderTemplateId", "RenderTemplateId"),
                SoType = SourceCode.SmartObjects.Services.ServiceSDK.Types.SoType.Text,
                Type = "System.String"
            };
            ContainerProperties.Add(rendertemplateid);

            Property pisearchresultid = new Property
            {
                Name = "pisearchresultid",
                MetaData = new MetaData("piSearchResultId", "piSearchResultId"),
                SoType = SourceCode.SmartObjects.Services.ServiceSDK.Types.SoType.Text,
                Type = "System.String"
            };
            ContainerProperties.Add(pisearchresultid);

            return ContainerProperties;

        }

        private List<Property> GetUserProperties()
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

            Property path = new Property
            {
                Name = "path",
                MetaData = new MetaData("Path", "Path"),
                SoType = SourceCode.SmartObjects.Services.ServiceSDK.Types.SoType.Text,
                Type = "System.String"
            };
            ContainerProperties.Add(path);

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

            Property serviceapplicationid = new Property
            {
                Name = "serviceapplicationid",
                MetaData = new MetaData("ServiceApplicationID", "ServiceApplicationID"),
                SoType = SourceCode.SmartObjects.Services.ServiceSDK.Types.SoType.Text,
                Type = "System.String"
            };
            ContainerProperties.Add(serviceapplicationid);

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

            Property docaclmeta = new Property
            {
                Name = "docaclmeta",
                MetaData = new MetaData("docaclmeta", "docaclmeta"),
                SoType = SourceCode.SmartObjects.Services.ServiceSDK.Types.SoType.Text,
                Type = "System.String"
            };
            ContainerProperties.Add(docaclmeta);

            Property originalpath = new Property
            {
                Name = "originalpath",
                MetaData = new MetaData("OriginalPath", "OriginalPath"),
                SoType = SourceCode.SmartObjects.Services.ServiceSDK.Types.SoType.Text,
                Type = "System.String"
            };
            ContainerProperties.Add(originalpath);

            Property resulttypeidlist = new Property
            {
                Name = "resulttypeidlist",
                MetaData = new MetaData("ResultTypeIdList", "ResultTypeIdList"),
                SoType = SourceCode.SmartObjects.Services.ServiceSDK.Types.SoType.Text,
                Type = "System.String"
            };
            ContainerProperties.Add(resulttypeidlist);

            Property partitionid = new Property
            {
                Name = "partitionid",
                MetaData = new MetaData("PartitionId", "PartitionId"),
                SoType = SourceCode.SmartObjects.Services.ServiceSDK.Types.SoType.Text,
                Type = "System.String"
            };
            ContainerProperties.Add(partitionid);

            Property urlzone = new Property
            {
                Name = "urlzone",
                MetaData = new MetaData("UrlZone", "UrlZone"),
                SoType = SourceCode.SmartObjects.Services.ServiceSDK.Types.SoType.Text,
                Type = "System.String"
            };
            ContainerProperties.Add(urlzone);

            Property aamenabledmanagedproperties = new Property
            {
                Name = "aamenabledmanagedproperties",
                MetaData = new MetaData("AAMEnabledManagedProperties", "AAMEnabledManagedProperties"),
                SoType = SourceCode.SmartObjects.Services.ServiceSDK.Types.SoType.Text,
                Type = "System.String"
            };
            ContainerProperties.Add(aamenabledmanagedproperties);

            Property resulttypeid = new Property
            {
                Name = "resulttypeid",
                MetaData = new MetaData("ResultTypeId", "ResultTypeId"),
                SoType = SourceCode.SmartObjects.Services.ServiceSDK.Types.SoType.Text,
                Type = "System.String"
            };
            ContainerProperties.Add(resulttypeid);

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

            Property rendertemplateid = new Property
            {
                Name = "rendertemplateid",
                MetaData = new MetaData("RenderTemplateId", "RenderTemplateId"),
                SoType = SourceCode.SmartObjects.Services.ServiceSDK.Types.SoType.Text,
                Type = "System.String"
            };
            ContainerProperties.Add(rendertemplateid);

            Property pisearchresultid = new Property
            {
                Name = "pisearchresultid",
                MetaData = new MetaData("piSearchResultId", "piSearchResultId"),
                SoType = SourceCode.SmartObjects.Services.ServiceSDK.Types.SoType.Text,
                Type = "System.String"
            };
            ContainerProperties.Add(pisearchresultid);


            return ContainerProperties;
        }

        private List<Property> GetSPSearchProperties()
        {
            List<Property> ContainerProperties = new List<Property>();

            ContainerProperties.AddRange(SPSearchProperties.GetSearchInputProperties());
            ContainerProperties.AddRange(SPSearchProperties.GetSearchResultSummaryProperties());
            ContainerProperties.AddRange(SPSearchProperties.GetSearchResultReturnProperties());
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
                        //returns.Where(p => p.Name.Equals("sort", StringComparison.OrdinalIgnoreCase)).First().Value = SerializedResults.Inputs.SortString;

                        // for testing
                        returns.Where(p => p.Name.Equals("sort", StringComparison.OrdinalIgnoreCase)).First().Value = GetColumns(SerializedResults);
                    }

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

            string searchsiteurl = string.Empty;
            var searchsiteurlprop = inputs.Where(p => p.Name.Equals("searchsiteurl", StringComparison.OrdinalIgnoreCase)).First();
            if (searchsiteurlprop != null && searchsiteurlprop.Value != null && !string.IsNullOrWhiteSpace(searchsiteurlprop.Value.ToString()))
            {
                searchsiteurl = searchsiteurlprop.Value.ToString();
                InputValues.SiteUrl = searchsiteurl;
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


            string fileext = string.Empty;
            var fileextprop = inputs.Where(p => p.Name.Equals("fileextensions", StringComparison.OrdinalIgnoreCase)).First();
            if (fileextprop != null && fileextprop.Value != null && !string.IsNullOrWhiteSpace(fileextprop.Value.ToString()))
            {
                InputValues.FileExtensions = new List<string>();
                fileext = fileextprop.Value.ToString();
                string[] sortsArray = fileext.Split(',');

                foreach(string fx in sortsArray)
                {
                    InputValues.FileExtensions.Add(fx.Trim());
                }

                string filter = string.Empty;
                for (int i = 0; i < InputValues.FileExtensions.Count; i++)
                {
                    filter += "\"" + InputValues.FileExtensions[i] + "\"";
                    if (i <= InputValues.FileExtensions.Count-2)
                    {
                        filter += ",";
                    }
                }
                InputValues.FileExtensionsString = filter;
            }


            string sorts = string.Empty;
            Dictionary<string, string> sort = new Dictionary<string, string>();
            var sortProp = inputs.Where(p => p.Name.Equals("sort", StringComparison.OrdinalIgnoreCase)).First();
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
                foreach(KeyValuePair<string, string> s in sort)
                {
                    InputValues.SortString += s.Key + ":" + s.Value;
                    if(o <= sort.Count-2)
                    {
                        InputValues.SortString += ",";
                    }
                    o++;
                }
            }

            bool enablestemming = false;
            var enablestemmingprop = inputs.Where(p => p.Name.Equals("enablestemming", StringComparison.OrdinalIgnoreCase)).First();
            if (enablestemmingprop != null && enablestemmingprop.Value != null && !string.IsNullOrWhiteSpace(enablestemmingprop.Value.ToString()))
            {
                if (bool.TryParse(enablestemmingprop.Value.ToString(), out enablestemming))
                {
                    InputValues.EnableStemming = enablestemming;
                }
            }

            bool trimduplicates = false;
            var trimduplicatesprop = inputs.Where(p => p.Name.Equals("trimduplicates", StringComparison.OrdinalIgnoreCase)).First();
            if (trimduplicatesprop != null && trimduplicatesprop.Value != null && !string.IsNullOrWhiteSpace(trimduplicatesprop.Value.ToString()))
            {
                if (bool.TryParse(trimduplicatesprop.Value.ToString(), out trimduplicates))
                {
                    InputValues.TrimDuplicates = trimduplicates;
                }
            }

            bool enablequeryrules = false;
            var enablequeryrulesprop = inputs.Where(p => p.Name.Equals("enablequeryrules", StringComparison.OrdinalIgnoreCase)).First();
            if (enablequeryrulesprop != null && enablequeryrulesprop.Value != null && !string.IsNullOrWhiteSpace(enablequeryrulesprop.Value.ToString()))
            {
                if (bool.TryParse(enablequeryrulesprop.Value.ToString(), out enablequeryrules))
                {
                    InputValues.EnableQueryRules = enablequeryrules;
                }
            }

            bool processbestbets = false;
            var processbestbetsprop = inputs.Where(p => p.Name.Equals("processbestbets", StringComparison.OrdinalIgnoreCase)).First();
            if (processbestbetsprop != null && processbestbetsprop.Value != null && !string.IsNullOrWhiteSpace(processbestbetsprop.Value.ToString()))
            {
                if (bool.TryParse(processbestbetsprop.Value.ToString(), out processbestbets))
                {
                    InputValues.ProcessBestBets = processbestbets;
                }
            }

            bool processpersonal = false;
            var processpersonalprop = inputs.Where(p => p.Name.Equals("processpersonal", StringComparison.OrdinalIgnoreCase)).First();
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
        private string GetColumns(RESTSearchResultsSerialized results)
        {
            string cols = string.Empty;
            int i = 0;
            foreach(ResultCell cell in results.SearchResults.Rows[0].Cells.OrderBy(p => p.Key))
            {
                cols += cell.Key + " (" + cell.ValueType + ")";
                if (i <= results.SearchResults.Rows.Length-2)
                {
                    cols += ",";
                }
                i++;
            }

            return cols;
        }

        private string BuildSearchText(SearchInputs Inputs)
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

            if (Inputs.FileExtensions != null && Inputs.FileExtensions.Count > 0)
            {
                //&refiners='filetype'

                if (Inputs.FileExtensions.Count < 2 )
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
