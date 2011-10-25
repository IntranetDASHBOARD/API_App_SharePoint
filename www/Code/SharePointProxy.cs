using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Xml;
using System.Xml.Linq;
using System.Text;
using System.Data;
using System.Net;
using System.IO;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace SharePointConnector
{
    
    public class SharePointProxy
    {
        private string _selectedSiteUrl;
        private string _originalSiteUrl;
        private NetworkCredential _sharepointCredentitals;
        private Lists _listWS;
        private Views _viewsWS;
        private Webs _websWS;

        /// <summary>
        /// Site Url of the item selected in SiteDetailTreeView
        /// </summary>
        public string SelectedSiteUrl
        {
            get { return _selectedSiteUrl; }
            set
            {
                if (!value.EndsWith("/"))
                {
                    _selectedSiteUrl = value + "/";
                }
                else
                {
                    _selectedSiteUrl = value;
                }
            }
        }

        /// <summary>
        /// Site Url specified in the Sharepoint Component component properties
        /// </summary>
        public string OriginalSiteUrl
        {
            get { return _originalSiteUrl; }
            set
            {
                if (!value.EndsWith("/"))
                {
                    _originalSiteUrl = value + "/";
                }
                else
                {
                    _originalSiteUrl = value;
                }
            }
        }

        /// <summary>
        /// Credential details used to connect to sharepoint web services
        /// </summary>
        public NetworkCredential SharepointCredentials
        {
            get { return _sharepointCredentitals; }
        }

        /// <summary>
        /// SharepointProxy class constructor
        /// </summary>
        /// <param name="siteUrl">Url of the sharepoint site to retrieve details from</param>
        /// <param name="sharepointCredentials">Credentials of the user accessing the sharepoint site</param>
        public SharePointProxy(string originalSiteUrl, NetworkCredential sharepointCredentials)
        {
            _originalSiteUrl = originalSiteUrl;
            _selectedSiteUrl = originalSiteUrl;
            _sharepointCredentitals = sharepointCredentials;
            InitialiseListWebserviceProperties();
            InitialiseViewsWebserviceProperties();
            InitialiseWebsWebserviceProperties();
        }


        #region Web service Initialization
        /// <summary>
        /// Initialize the List Web Service
        /// </summary>
        private void InitialiseListWebserviceProperties()
        {
            _listWS = new Lists(GetHostName());
            //_listWS = new Lists.Lists();
            _listWS.PreAuthenticate = true;
            _listWS.Credentials = _sharepointCredentitals;
            if (!_originalSiteUrl.EndsWith("/"))
            {
                _originalSiteUrl = _originalSiteUrl + "/";
            }
            _listWS.Url = _originalSiteUrl + "_vti_bin/lists.asmx";
        }

        /// <summary>
        /// Initialize the Views Web Service
        /// </summary>
        private void InitialiseViewsWebserviceProperties()
        {
            //_viewsWS = new Views.Views();
            _viewsWS = new Views(GetHostName());
            _viewsWS.PreAuthenticate = true;
            _viewsWS.Credentials = _sharepointCredentitals;
            //_viewsWS.Credentials = new NetworkCredential("thushari.desilva", "Clayton123", "adweb");
            if (!_originalSiteUrl.EndsWith("/"))
            {
                _originalSiteUrl = _originalSiteUrl + "/";
            }
            _viewsWS.Url = _originalSiteUrl + "_vti_bin/Views.asmx";
        }

        /// <summary>
        /// Initialize the Webs Web Service
        /// </summary>
        private void InitialiseWebsWebserviceProperties()
        {
            //_websWS = new Webs.Webs();
            _websWS = new Webs(GetHostName());
            _websWS.PreAuthenticate = true;
            _websWS.Credentials = _sharepointCredentitals;
            //_websWS.Credentials = new NetworkCredential("thushari.desilva", "Clayton123", "adweb");
            if (!_originalSiteUrl.EndsWith("/"))
            {
                _originalSiteUrl = _originalSiteUrl + "/";
            }
            _websWS.Url = _originalSiteUrl + "_vti_bin/webs.asmx";
        }
        #endregion


        /// <summary>
        /// Returns the Title and the Url of the Root Sharepoint site based on the Url provided by the user in the component properties
        /// </summary>
        /// <param name="siteTitle">Tile of the site</param>
        /// <param name="url">Url of the site</param>
        public void GetRootSiteDetails(out string siteTitle, out string url)
        {
            siteTitle = string.Empty;
            url = string.Empty;
            _websWS.Url = _originalSiteUrl + "_vti_bin/webs.asmx";
            XElement site = XElement.Parse(_websWS.GetWeb(_originalSiteUrl.TrimEnd(("/").ToCharArray())).OuterXml);
            siteTitle = site.Attribute("Title").Value;
            url = site.Attribute("Url").Value;
        }


        /// <summary>
        /// Returns any sites under the Root sharepoint site
        /// </summary>
        /// <returns>List of sites under the specified site</returns>
        public IEnumerable<KeyValuePair<string, string>> GetSiteCollection()
        {
            _websWS.Url = _selectedSiteUrl + "_vti_bin/webs.asmx";
            XElement siteCollection = XElement.Parse(_websWS.GetWebCollection().OuterXml);
            IEnumerable<KeyValuePair<string, string>> siteList = from sites in siteCollection.Elements() 
                                                                 select new KeyValuePair<string, string>(sites.Attribute("Title").Value, sites.Attribute("Url").Value);
            return siteList;
        }


        /// <summary>
        /// Get all List items that are of the specified List Type 
        /// </summary>
        /// <param name="listType">TemplateId of the requried List Type</param>
        /// <returns>Collection of list items belonging to the specified list type</returns>
        public List<SharePointItem> GetListCollection(string listType)
        {
            _listWS.Url = _selectedSiteUrl + "_vti_bin/lists.asmx";
            List<SharePointItem> listCollection = new List<SharePointItem>();
            XElement listCollectionNode = XElement.Parse(_listWS.GetListCollection().OuterXml);

            IEnumerable<XElement> listTypeItems = from listItem in listCollectionNode.Elements()
                                                  where listItem.Attribute("ServerTemplate").Value == listType
                                                  select listItem;


            foreach (XElement element in listTypeItems)
            {
                SharePointItem list = new SharePointItem();
                list.Guid = element.Attribute("ID").Value;
                list.Title = element.Attribute("Title").Value;
                listCollection.Add(list);
            }
            return listCollection;
        }

        /// <summary>
        /// Get all Non Hidden Views that belong to a List
        /// </summary>
        /// <param name="selectedListGuid">GUID of the selected list in the tree view</param>
        /// <returns></returns>
        public List<SharePointItem> GetSharepointViewCollection(string selectedListGuid)
        {
            _viewsWS.Url = _selectedSiteUrl + "/_vti_bin/Views.asmx";
            List<SharePointItem> viewCollection = new List<SharePointItem>();
            XElement allViewElements = XElement.Parse(_viewsWS.GetViewCollection(selectedListGuid).OuterXml);
            IEnumerable<XElement> viewElements = from view in allViewElements.Elements() 
                                                 where view.Attributes("Hidden").Any() == false 
                                                 select view;

            foreach (XElement element in viewElements)
            {
                SharePointItem view = new SharePointItem();
                view.Guid = element.Attribute("Name").Value;
                view.Title = element.Attribute("DisplayName").Value;
                viewCollection.Add(view);
            }
            return viewCollection;
        }


        /// <summary>
        /// Return a list of items that belong to a selected list and a view
        /// </summary>
        /// <param name="listId">GUID of the selected list</param>
        /// <param name="viewName">GUID of the selected views </param>
        /// <param name="url">Url of the selected site</param>
        /// <param name="linkDetails">Name and the column position of the hyperlink column</param>
        /// <returns></returns>
        public DataTable GetItemList(string viewName, string listId, string url, Guid componentKey, out string linkDetails)
        {
            linkDetails = string.Empty;
            string hostName = GetHostName();

            IEnumerable<XElement> rowElements = GetItems(listId, viewName, url);

            IEnumerable<KeyValuePair<string, SharepointField>> viewFields = LoadSharepointViews(viewName, listId);

            DataTable tempTable = new DataTable();
            DataColumn tempColumn;
            tempTable.Columns.Add("ItemID");
            DataRow tempRow;
            int count = 2;
            foreach (KeyValuePair<string, SharepointField> displayName in viewFields)
            {
                if (displayName.Key.Equals("LinkFilename") || displayName.Key.Equals("LinkTitle"))
                {
                    //return details to the calling method to enable the LinkFilename column to be made a link button
                    linkDetails = displayName.Value.DisplayName + "|" + count;
                }
                tempColumn = new DataColumn();
                tempColumn.DataType = Type.GetType("System.String");
                tempColumn.ColumnName = displayName.Value.DisplayName; 
                tempTable.Columns.Add(tempColumn);
                count++;
            }

            if (viewFields.Count() > 0)
            {
                foreach (XElement row in rowElements)
                {
                    tempRow = tempTable.NewRow();
                    foreach (KeyValuePair<string, SharepointField> field in viewFields)
                    {
                        if (row.Attribute("ows_" + field.Key) != null)
                        {
                            if (field.Key.Equals("DocIcon"))
                            {
                                tempRow[field.Value.DisplayName] = "<img src='../GetImage.aspx?key=" + HttpUtility.UrlEncode(componentKey.ToString()) + "&host=" + HttpUtility.UrlEncode(hostName) + "&fileType=" + HttpUtility.UrlEncode(GetDocIconFileName(row.Attribute("ows_" + field.Key).Value)) + "'/>";
                            }
                            else if (field.Key.Equals("Attachments"))
                            {
                                if (Convert.ToInt32(row.Attribute("ows_" + field.Key).Value) == 1)
                                {
                                    tempRow[field.Value.DisplayName] = "<img src='../GetImage.aspx?key=" + HttpUtility.UrlEncode(componentKey.ToString()) + "&host=" + HttpUtility.UrlEncode(hostName) + "&fileType=" + HttpUtility.UrlEncode("attach.gif") +"'/>";
                                }
                                else
                                {
                                    tempRow[field.Value.DisplayName] = string.Empty;
                                }
                            }
                            else
                            {
                                tempRow[field.Value.DisplayName] = FormatValue(row.Attribute("ows_" + field.Key).Value, field.Value.IsPercentage);
                            }
                        }
                        else
                        {
                            if (row.Attribute("ows_DocIcon") == null & row.Attribute("ows_ContentType").Value.ToLower().Equals("folder"))
                            {
                                tempRow[field.Value.DisplayName] = "<img src='../GetImage.aspx?key=" + HttpUtility.UrlEncode(componentKey.ToString()) + "&host=" + HttpUtility.UrlEncode(hostName) + "&fileType=" + HttpUtility.UrlEncode("folder.gif") + "'/>";
                            }
                        }
                        //ContentType|Url|ContentTypeId -- required when retrieving document or folder details in display mode
                        tempRow["ItemID"] = row.Attribute("ows_ContentType").Value + "|" + row.Attribute("ows_ServerUrl").Value + "|" + row.Attribute("ows_ContentTypeId").Value; 
                    }
                    tempTable.Rows.Add(tempRow);
                }
            }
            return tempTable;
        }


        /// <summary>
        /// Get the image related to a particular File Type
        /// </summary>
        /// <param name="fileType">file type extention</param>
        /// <returns>name of the file from the DocIcon.xml file</returns>
        private string GetDocIconFileName(string fileType)
        {
            string fileName = string.Empty;
            XElement docIconElements = XElement.Load(AppDomain.CurrentDomain.BaseDirectory + "/Includes/DOCICON.XML");

            if ((from iconElement in docIconElements.Descendants("Mapping").Attributes("Key") where iconElement.Value == fileType select iconElement).Any())
            {
                fileName = (from iconElement in docIconElements.Descendants("Mapping") where iconElement.Attribute("Key").Value == fileType select iconElement).First().Attribute("Value").Value.ToString();
            }
            else
            {
                fileName = "icgen.gif";
            }
            return fileName;
        }


        /// <summary>
        ///  Return a list of views that belongs to a list 
        /// </summary>
        /// <param name="viewId">GUID of the selected View</param>
        /// <param name="listName">Name of the selected List</param>
        /// <param name="defaultSortOrder"></param>
        /// <returns>Key Value pairs which contains the Name and the Display Name of views</returns>
        private IEnumerable<KeyValuePair<string, SharepointField>> LoadSharepointViews(string viewId, string listName)
        {
            _listWS.Url = _selectedSiteUrl + "_vti_bin/lists.asmx";
            XElement listAndViewElement = XElement.Parse(_listWS.GetListAndView(listName, viewId).OuterXml);
            XNamespace nameSpace = listAndViewElement.Name.Namespace;

            IEnumerable<KeyValuePair<string, SharepointField>> displayFields = from viewFields in listAndViewElement.Descendants(nameSpace + "ViewFields").Elements(nameSpace + "FieldRef")
                                                                      from listNodes in listAndViewElement.Descendants(nameSpace + "Field")
                                                                      where listNodes.Attribute("Name").Value == viewFields.Attribute("Name").Value
                                                                      select new KeyValuePair<string, SharepointField>
                                                                      (
                                                                          (string)listNodes.Attribute("Name"),
                                                                          //(string)listNodes.Attribute("DisplayName")
                                                                          new SharepointField
                                                                              (
                                                                                listNodes.Attribute("DisplayName").Value, 
                                                                                listNodes.Attribute("Type").Value,
                                                                                listNodes.Attributes("Percentage").Any() ? Convert.ToBoolean(listNodes.Attribute("Percentage").Value) : false
                                                                              )
                                                                      );
            return displayFields;

        }


        /// <summary>
        /// 
        /// </summary>
        /// <param name="listName">Name of the selected List</param>
        /// <param name="viewName">GUID of the selected View</param>
        /// <param name="url">Url of the selected item</param>
        /// <returns></returns>
        private IEnumerable<XElement> GetItems(string listId, string viewId, string url)
        {
            _listWS.Url = _selectedSiteUrl + "_vti_bin/lists.asmx";
            int rowLimit = 100;
            XmlDocument xmlDoc = new XmlDocument();
            XmlNode nodeQuery = xmlDoc.CreateNode(XmlNodeType.Element, "Query", "");
            XmlNode nodeViewFields = xmlDoc.CreateNode(XmlNodeType.Element, "ViewFields", "");
            XmlNode nodeQueryOptions = xmlDoc.CreateNode(XmlNodeType.Element, "QueryOptions", "");

            if (!url.Equals(string.Empty))
            {
                AddXmlElement(ref nodeQueryOptions, "<Folder>value</Folder>".Replace("value", url));
            }

            XElement listItems = XElement.Parse(_listWS.GetListItems(listId, viewId, nodeQuery, nodeViewFields, rowLimit.ToString(), nodeQueryOptions, null).OuterXml);

            IEnumerable<XElement> rowElements = from rows in listItems.Descendants()
                                                where rows.Name.LocalName.ToLower() == "row"
                                                select rows;
            return rowElements;
        }


        private void AddXmlElement(ref XmlNode elementNode, string element)
        {
            elementNode.InnerXml += element;
        }


        /// <summary>
        /// Download the selected Sharepoint File
        /// </summary>
        /// <param name="itemUrl">Url of the selected file</param>
        /// <param name="contentType">Content Type of the file</param>
        /// <returns></returns>
        public byte[] DownloadFile(string itemUrl, out string contentType)
        {
            string hostName = GetHostName();
            WebClient sharepointClient = new WebClient();
            sharepointClient.Credentials = _sharepointCredentitals;
            byte[] content = sharepointClient.DownloadData(hostName + itemUrl);
            contentType = sharepointClient.ResponseHeaders.Get("Content-Type");
            return content;
        }


        /// <summary>
        /// 
        /// </summary>
        /// <param name="listId">GUID of the selected item</param>
        /// <param name="viewId">GUID of the selected view</param>
        /// <param name="contentTypeId">Id of the selected Content Type</param>
        /// <param name="itemUrl">Url of the selected item</param>
        /// <returns></returns>
        public IEnumerable<KeyValuePair<string, string>> GetTaskDetails(string listId, string viewId, string contentTypeId, string itemUrl)
        {
            XElement contentTypeElement = XElement.Parse(_listWS.GetListContentType(listId, contentTypeId).OuterXml);

            IEnumerable<XElement> rowElements = GetItems(listId, viewId, string.Empty);

            XElement selectedItem = (from item in rowElements
                                     where item.Attribute("ows_ServerUrl").Value == itemUrl
                                     select item).First();

            IEnumerable<KeyValuePair<string, string>> taskDisplayDetails = from XElement fields in contentTypeElement.Descendants(XName.Get("Field", contentTypeElement.Name.Namespace.ToString()))
                                                                           where fields.Attributes("Group").Any() == false
                                                                           select new KeyValuePair<string, string>
                                                                           (
                                                                               "<b>" + fields.Attribute("DisplayName").Value + "</b>",
                                                                               selectedItem.Attributes("ows_" + fields.Attribute("Name").Value).Any() ?
                                                                                    FormatValue(selectedItem.Attribute("ows_" + fields.Attribute("Name").Value).Value, fields.Attributes("Percentage").Any() ? 
                                                                                        Convert.ToBoolean(fields.Attribute("Percentage").Value) : false) 
                                                                                    : string.Empty
                                                                           );

            IEnumerable<KeyValuePair<string, string>> attachmentList = GetAttachmentCollection(listId, selectedItem.Attribute("ows_ID").Value);
            taskDisplayDetails = taskDisplayDetails.Concat(attachmentList);
            return taskDisplayDetails;
        }


        /// <summary>
        /// Get attachments of a selected item
        /// </summary>
        /// <param name="listId">GUID of the selected item</param>
        /// <param name="listItemId">ID of the the selected item</param>
        /// <returns></returns>
        private IEnumerable<KeyValuePair<string, string>> GetAttachmentCollection(string listId, string listItemId)
        {
            XElement attachements = XElement.Parse(_listWS.GetAttachmentCollection(listId, listItemId).OuterXml);
            IEnumerable<KeyValuePair<string, string>> attachmentList = from attachement in attachements.Descendants(XName.Get("Attachment", attachements.Name.Namespace.ToString()))
                                                                       select new KeyValuePair<string, string>
                                                                       (
                                                                            "<b>Attachments</b>",
                                                                            "<a href='" + attachement.Value + "'target='_blank'>" + attachement.Value.Substring(attachement.Value.LastIndexOf('/') + 1) + "</a>"
                                                                       );
            return attachmentList;
        }


        /// <summary>
        /// Remove ID's from the display value and format percentage values
        /// </summary>
        /// <param name="value">Vlaue to be formatted</param>
        /// <returns>Formatted value</returns>
        private string FormatValue(string value, bool percentage)
        {
            int startIndex = 0;
            if (value.IndexOf(";#") > -1)
            {
                startIndex = value.IndexOf(";#") + 2;
                value = value.Substring(startIndex, value.Length - startIndex);
            }
            if (percentage)
            {
                double tempValue;
                double.TryParse(value, out tempValue);
                value = tempValue.ToString("0%");
            }
            return value;
        }

        /// <summary>
        /// Return the hostname for the specified site
        /// </summary>
        /// <returns>hostname</returns>
        public string GetHostName()
        {
            if (OriginalSiteUrl.StartsWith("http://"))
            {
                return ("http://" + new Uri(OriginalSiteUrl).Host);
            }
            else if(OriginalSiteUrl.StartsWith("https://"))
            {
                return ("https://" + new Uri(OriginalSiteUrl).Host);
            }
            return string.Empty;
        }
    }



    public class SharePointItem
    {
        private string _guid;
        private string _title;


        public string Guid
        {
            get { return _guid; }
            set { _guid = value; }
        }

        public string Title
        {
            get { return _title; }
            set { _title = value; }
        }

    }

    public class SharepointField
    {
        private string _displayName;
        private string _fieldType;
        private bool _isPercentage;

        public SharepointField(string displayName, string fieldType, bool isPerecentage)
        {
            _displayName = displayName;
            _fieldType = fieldType;
            _isPercentage = isPerecentage;
        }

        public string DisplayName
        {
            get { return _displayName; }
            set { _displayName = value; }
        }

        public string FieldType
        {
            get { return _fieldType; }
            set { _fieldType = value; }
        }

        public bool IsPercentage
        {
            get { return _isPercentage; }
            set { _isPercentage = value; }
        }
    }

}
