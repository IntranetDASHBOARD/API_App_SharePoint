using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.Script.Serialization;
using System.Xml.Linq;
using System.Net;
using System.Text;
using System.Xml;
using System.IO;
using System.Web.Configuration;
using System.Configuration;
using System.Data;

using IntranetDASHBOARD.API;


namespace SharePointConnector
{
    public partial class Default : iDCMSComponent
    {
        private string _siteUrl;
        private string _username;
        private string _password;
        private string _domain;
        private string _linkColumnDetails;
        private bool _expandUsingWS = true;
        
        private SharePointConnector.SharePointProxy _Sharepoint;

        //Cms component field initialization
        public iDCMSComponentField selectedList = new iDCMSComponentField(string.Empty);
        public iDCMSComponentField selectedView = new iDCMSComponentField(string.Empty);
        public iDCMSComponentField selectedSiteUrl = new iDCMSComponentField(string.Empty);
        public iDCMSComponentField selectedNode = new iDCMSComponentField(string.Empty);
        public iDCMSComponentField propertyDetails = new iDCMSComponentField(string.Empty);

        
        /// <summary>
        /// Create SharepointProxy Object, based on the values from component properties 
        /// </summary>
        private void CreateSharepointProxyObject()
        {
            if (_Sharepoint == null)
            {
                //retrieve information from component properties
                _siteUrl = ValidSiteUrl(GetExternalComponentPropertyValue("SiteUrl"));
                _username = GetExternalComponentPropertyValue("Username");
                _password = GetExternalComponentPropertyValue("Password");
                _domain = GetExternalComponentPropertyValue("Domain");

                NetworkCredential sharepointCredentials = new NetworkCredential(_username, _password, _domain);
                _Sharepoint = new SharePointProxy(_siteUrl, sharepointCredentials);
            }
        }        

        
        protected override void OnLoad(EventArgs e)
        {
            if (!IsPostBack)
            {
                base.BindCMSComponentFieldToControl(selectedListTextBox, selectedList.FormInputName);
                base.BindCMSComponentFieldToControl(selectedViewTextBox, selectedView.FormInputName);
                base.BindCMSComponentFieldToControl(selectedSiteUrlTextBox, selectedSiteUrl.FormInputName);
                base.BindCMSComponentFieldToControl(selectedNodePathTextBox, selectedNode.FormInputName);
                base.BindCMSComponentFieldToControl(propertyDetailsTextBox, propertyDetails.FormInputName);
                base.OnLoad(e);
            }
        }



#region EditMode 

        /// <summary>
        /// Initial method called in Edit Mode
        /// </summary>
        protected override void OnLoadOfEditMode()
        {
            
            if (!GetExternalComponentPropertyValue("SiteUrl").Equals(string.Empty) && !GetExternalComponentPropertyValue("Username").Equals(string.Empty))
            {
                //save current component property details in a CmsComponentField so that it could be compared in display mode
                propertyDetailsTextBox.Text = GetExternalComponentPropertyValue("SiteUrl") + "|" + GetExternalComponentPropertyValue("Username");

                
                //persist or remove selected tree node details 
                if (CheckIfComponentPropertiesUpdated())
                {
                    //if component properties have been updated clear persisted selected tree node values
                    ClearSelectedTreeViewValues();
                    Session.Remove(FormPrefix + "treeViewXml");
                }
                else
                {
                    PersistSelectedTreeViewValues();
                }
                

                //load tree view using the session object or sharepoint web services
                if (Session[FormPrefix + "treeViewXml"] != null)
                {
                    DeserializeTreeView();
                }
                else
                {
                    if (ValidateComponentProperties(Mode.Edit))
                    {
                        CreateSharepointProxyObject();
                        LoadSiteDetailTreeView();
                    }
                }
                SelectNode();
            }
            else
            {
                editModeErrorPanel.Visible = true;
                editModeErrorTitle.Text = "SharePoint URL";
                editModeErrorMessage.Text = "Please select the SiteUrl and Username for this component by using the Properties functionality located in the Component menu.";
            }
        }


        /// <summary>
        /// Create the root node of the sharepoint tree based on the Root Sharepoint Url provided by the user in the component properties
        /// </summary>
        private void LoadSiteDetailTreeView()
        {
            if (_Sharepoint != null)
            {
                //clear any existing nodes of the tree view
                SiteDetailTreeView.Nodes.Clear();
                try
                {
                    string siteTitle, siteUrl = string.Empty;
                    _Sharepoint.GetRootSiteDetails(out siteTitle, out siteUrl);
                    if (!siteTitle.Equals(string.Empty) && !siteUrl.Equals(string.Empty))
                    {
                        TreeNode rootNode = new TreeNode();
                        rootNode.Text = siteTitle;
                        rootNode.Value = "Site|" + siteUrl;
                        rootNode.PopulateOnDemand = true;
                        rootNode.SelectAction = TreeNodeSelectAction.None;
                        SiteDetailTreeView.Nodes.Add(rootNode);
                    }
                    else
                    {
                        editModeErrorPanel.Visible = true;
                        editModeErrorMessage.Text = "There are no items in the list.";
                    }
                    ExpandTreeView();
                }
                catch (System.Net.WebException webEx)
                {
                    HttpWebResponse response = (HttpWebResponse)webEx.Response;
                    if (response.StatusCode == HttpStatusCode.Unauthorized)
                    {
                        editModeErrorPanel.Visible = true;
                        editModeErrorTitle.Text = "Invalid Credentials";
                        editModeErrorMessage.Text = "You are not authorized to access this SharePoint content at this time. Please contact the Page Owner for more information.";
                    }
                    else if (response.StatusCode == HttpStatusCode.NotFound)
                    {
                        editModeErrorPanel.Visible = true;
                        editModeErrorTitle.Text = "Invalid URL";
                        editModeErrorMessage.Text = "The URL specified within the control is invalid, please update the SharePoint details for this component by using the Properties functionality located in the Component menu. Alternatively contact the Page Owner for more information.";
                    }
                }
                catch (Exception ex)
                {
                    editModeErrorMessage.Text = ex.Message;
                    editModeErrorPanel.Visible = true;
                }
            }
        }


        /// <summary>
        /// Expand all the appropriate tree nodes
        /// </summary>
        private void ExpandTreeView()
        {
            try
            {
                string selectedNodePath = string.Empty;
                string[] selectedNodes = selectedNodePathTextBox.Text.Split('!');
                if (selectedNodes.Length > 0)
                {
                    for (int i = 0; i < selectedNodes.Length; i++)
                    {
                        if (SiteDetailTreeView.FindNode(selectedNodePath + selectedNodes[i]) != null)
                        {
                            if (i < selectedNodes.Length - 1)
                            {
                                SiteDetailTreeView.FindNode(selectedNodePath + selectedNodes[i]).Expand();
                                selectedNodePath = selectedNodePath + selectedNodes[i] + "!";
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                editModeErrorPanel.Visible = true;
                editModeErrorMessage.Text = ex.Message;
            }
        }

        /// <summary>
        /// Select the previously selected treeview node
        /// </summary>
        private void SelectNode()
        {
            try
            {
                if (selectedNodePathTextBox.Text.Trim().Length > 0)
                {
                    TreeNode selectedNode = SiteDetailTreeView.FindNode(selectedNodePathTextBox.Text);
                    if (selectedNode != null)
                    {
                        selectedNode.Select();
                        scriptLiteral.Text = "<script>document.getElementById('" + selectedViewTextBox.Text + "').checked = true;</script>";
                    }
                }
            }
            catch (Exception ex)
            {
                editModeErrorMessage.Text = ex.Message;
                editModeErrorPanel.Visible = true;
            }
        }


        /// <summary>
        /// Method called when a node is expanded to create the relevant child nodes in the tree view
        /// </summary>
        public void SiteDetailTreeView_Expanded(object s, TreeNodeEventArgs e)
        {
            if (_expandUsingWS == true && e.Node.ChildNodes.Count == 0)
            {
                CreateSharepointProxyObject();

                //e.Node.ChildNodes.Clear();
                TreeNode childNode;
                try
                {
                    if (e.Node.Value.StartsWith("Site"))
                    {
                        if (e.Node.Value.Split('|').Count() > 1)
                        {
                            _Sharepoint.SelectedSiteUrl = e.Node.Value.Split('|')[1];
                        }
                        IEnumerable<KeyValuePair<string, string>> siteList = _Sharepoint.GetSiteCollection();
                        if (siteList.Count() > 0)
                        {
                            foreach (KeyValuePair<string, string> site in siteList)
                            {
                                childNode = new TreeNode();
                                childNode.Text = site.Key;
                                childNode.Value = "Site|" + site.Value;
                                childNode.PopulateOnDemand = true;
                                childNode.SelectAction = TreeNodeSelectAction.None;
                                e.Node.ChildNodes.Add(childNode);
                            }
                        }

                        SharePointConnector.SharePointSettings sharepointSettings = (SharePointSettings)System.Configuration.ConfigurationManager.GetSection("SharePointSettings");
                        if (sharepointSettings != null)
                        {
                            for (int i = 0; i < sharepointSettings.SharepointLists.Count; i++)
                            {
                                childNode = new TreeNode();
                                childNode.Text = sharepointSettings.SharepointLists[i].Name;
                                childNode.Value = "ListType|" + sharepointSettings.SharepointLists[i].TemplateId;
                                childNode.PopulateOnDemand = true;
                                childNode.SelectAction = TreeNodeSelectAction.None;
                                e.Node.ChildNodes.Add(childNode);
                            }
                        }
                    }
                    else if (e.Node.Value.StartsWith("ListType"))
                    {
                        string listType = e.Node.Value.Split('|')[1];
                        _Sharepoint.SelectedSiteUrl = e.Node.Parent.Value.Split('|')[1];

                        List<SharePointConnector.SharePointItem> listCollection = _Sharepoint.GetListCollection(listType);
                        foreach (SharePointConnector.SharePointItem listItem in listCollection)
                        {
                            childNode = new TreeNode();
                            childNode.Text = listItem.Title;
                            childNode.Value = "List|" + listItem.Guid + "|" + listItem.Title;
                            childNode.PopulateOnDemand = true;
                            childNode.SelectAction = TreeNodeSelectAction.None;
                            e.Node.ChildNodes.Add(childNode);
                        }
                    }
                    else if (e.Node.Value.StartsWith("List"))
                    {
                        string selectedListGuid = e.Node.Value.Split('|')[1];
                        _Sharepoint.SelectedSiteUrl = e.Node.Parent.Parent.Value.Split('|')[1];
                        List<SharePointConnector.SharePointItem> viewCollection = _Sharepoint.GetSharepointViewCollection(selectedListGuid);
                        foreach (SharePointConnector.SharePointItem viewItem in viewCollection)
                        {

                            childNode = new TreeNode();
                            childNode.Text = "<input type='radio' name='sharepointTreeNode' id='" + viewItem.Guid + "' onClick='javascript: radioButtonClicked(this);'/>" + viewItem.Title;
                            childNode.Value = "View|" + viewItem.Guid;
                            childNode.SelectAction = TreeNodeSelectAction.Select;
                            e.Node.ChildNodes.Add(childNode);
                        }
                    }
                    editModeErrorPanel.Visible = false;
                }
                catch (Exception ex)
                {
                    editModeErrorPanel.Visible = true;
                    editModeErrorMessage.Text = ex.Message;
                }
            }

        }

        /// <summary>
        /// Save ListGuid, ViewName and the SiteUrl of the selected view, to enable details to be used in display mode when loading Sharepoint Items
        /// Save the path of the selected node to be used in edit mode
        /// </summary>
        public void SiteDetailTreeView_NodeSelected(object s, EventArgs e)
        {
            if (SiteDetailTreeView.SelectedNode.Value.StartsWith("View|"))
            {
                if (SiteDetailTreeView.SelectedNode.Parent.Value.Split('|').Length == 3) 
                {
                    selectedNodePathTextBox.Text = SiteDetailTreeView.SelectedNode.ValuePath; //Value path of the selected node, to enable the correct node to be selected when sweitching between edit and display modes
                    selectedListTextBox.Text = SiteDetailTreeView.SelectedNode.Parent.Value.Split('|')[1]; ; //GUID of the selected List
                    selectedViewTextBox.Text = SiteDetailTreeView.SelectedNode.Value.Split('|')[1]; //GUID of the selected View
                    selectedSiteUrlTextBox.Text = SiteDetailTreeView.SelectedNode.Parent.Parent.Parent.Value.Split('|')[1]; //site url of the selected site
                    SerializeTreeView();
                    scriptLiteral.Text = "<script>document.getElementById('" + selectedViewTextBox.Text + "').checked = true;</script>";
                }
            }
        }


        /// <summary>
        /// To add the treeview to session, create a xml document
        /// </summary>
        private void SerializeTreeView()
        {
            XElement treeViewXmlDoc = new XElement("TreeView");
            SaveTreeNodes(SiteDetailTreeView.Nodes, treeViewXmlDoc);
            Session.Add(FormPrefix + "treeViewXml", treeViewXmlDoc.ToString());
        }


        /// <summary>
        /// Create xml elements for child nodes and add them to there parent element in the xml document
        /// </summary>
        /// <param name="nodeCollection">Collection of child nodes</param>
        /// <param name="parentNode">Parent element that the node collection belongs to</param>
        private void SaveTreeNodes(TreeNodeCollection nodeCollection, XElement parentNode)
        {
            XElement treeNode;
            foreach (TreeNode node in nodeCollection)
            {
                treeNode = new XElement("Node");
                treeNode.SetAttributeValue("Text", node.Text);
                treeNode.SetAttributeValue("Value", node.Value);
                parentNode.Add(treeNode);
                if (node.ChildNodes.Count > 0 && node.Expanded == true)
                {
                    SaveTreeNodes(node.ChildNodes, treeNode);
                }
            }
        }

        /// <summary>
        /// Using the tree view session object create tree nodes and add to SiteDetailTreeView
        /// </summary>
        private void DeserializeTreeView()
        {
            _expandUsingWS = false;

            XElement element = (XElement.Parse(Session[FormPrefix + "treeViewXml"].ToString())).Descendants().First();
            TreeNode rootNode = new TreeNode();
            rootNode.Text = element.Attribute("Text").Value;
            rootNode.Value = element.Attribute("Value").Value;
            rootNode.PopulateOnDemand = true;
            rootNode.SelectAction = TreeNodeSelectAction.None;
            SiteDetailTreeView.Nodes.Add(rootNode);
            rootNode.Expand();
            if (element.HasElements)
            {
                CreateTreeNodes(element, rootNode);
            }
            
            
        }

        /// <summary>
        /// Create child nodes and add the created nodes to the parent node
        /// </summary>
        /// <param name="parentElement">Xml for the parent element, which includes xml for any child elements</param>
        /// <param name="parentNode">parent tree view node, that child tree nodes needs to be added to</param>
        private void CreateTreeNodes(XElement parentElement, TreeNode parentNode)
        {
            TreeNode node;
            foreach (XElement childElement in parentElement.Elements())
            {
                node = new TreeNode();
                node.Text = childElement.Attribute("Text").Value;
                node.Value = childElement.Attribute("Value").Value;
                

                if (node.Value.StartsWith("View|"))
                {
                    node.SelectAction = TreeNodeSelectAction.Select;
                }
                else
                {
                    node.SelectAction = TreeNodeSelectAction.None;
                    node.PopulateOnDemand = true;
                }

                parentNode.ChildNodes.Add(node);

                if (childElement.HasElements)
                {
                    node.Expand();
                    CreateTreeNodes(childElement, node);
                }
            }
        }

#endregion


#region DisplayMode

        /// <summary>
        /// Inititial method called in Display Mode
        /// </summary>
        protected override void OnLoadOfDisplayMode()
        {
            if (!GetExternalComponentPropertyValue("SiteUrl").Equals(string.Empty) && !GetExternalComponentPropertyValue("Username").Equals(string.Empty))
            {
                if (!CheckIfComponentPropertiesUpdated())
                {
                    CreateSharepointProxyObject();
                    _Sharepoint.SelectedSiteUrl = selectedSiteUrl.Value;
                    LoadSharepointItems(string.Empty);
                }
                else // display message to select new values from tree view as component property values has been updated
                {
                    if (ValidateComponentProperties(Mode.Display)) 
                    {
                        displayModeErrorPanel.Visible = true;
                        displayModeErrorTitile.Text = "Select Source";
                        displayModeErrorMessage.Text = "Please select the SharePoint information source you wish to access by using the Edit Content functionality located in the Component menu";
                        //removed edit mode tree view from session as the component properties have changed and the tree view must be re-created
                        Session.Remove(FormPrefix + "treeViewXml");
                        Session.Remove(FormPrefix + "SharepointListData");
                    }
                }
            }
            else
            {
                displayModeErrorPanel.Visible = true;
                displayModeErrorTitile.Text = "SharePoint URL";
                displayModeErrorMessage.Text = "Please select the SiteUrl and Username for this component by using the Properties functionality located in the Component menu.";
            }
        }

        /// <summary>
        /// Compare previously entered component property values with the current component property values. 
        /// Prompt user to reselect details from the tree view if values have changed, so that sharepoint items will be retrieved using the updated component property details.
        /// </summary>
        /// <returns>boolean value indicating if component properties have been updated</returns>
        private bool CheckIfComponentPropertiesUpdated()
        {
            string siteUrl = string.Empty;
            string username = string.Empty;

            if (!propertyDetails.Value.Equals(string.Empty))
            {
                siteUrl = propertyDetails.Value.Split('|')[0];
                username = propertyDetails.Value.Split('|')[1];
            }

            if(!siteUrl.Equals(string.Empty) && !username.Equals(string.Empty))
            {
                if (GetExternalComponentPropertyValue("SiteUrl").Equals(siteUrl) && GetExternalComponentPropertyValue("Username").Equals(username))
                {
                    return false;
                }
                else
                {
                    
                    return true;
                }
            }
            else
            {
                return true;
            }
            
        }

        /// <summary>
        /// retrieve sharepoint data using sharepoint web services
        /// </summary>
        /// <param name="itemUrl">Server Url of the selected site or item</param>
        /// <param name="useSessionData">Boolean value to determine if data should be retrieved from sharepoint web services or session object</param>
        /// <returns></returns>
        private DataTable RetrieveSharepointListData(string itemUrl, bool useSessionData)
        {
            DataTable sharepointListData = new DataTable();
            try
            {
                if (!selectedList.Value.Equals(string.Empty) && !selectedView.Value.Equals(string.Empty))
                {
                    if (useSessionData && Session[FormPrefix + "SharepointListData"] != null)
                    {
                        sharepointListData = (DataTable)Session[FormPrefix + "SharepointListData"];
                    }
                    else
                    {
                        CreateSharepointProxyObject();
                        sharepointListData = _Sharepoint.GetItemList(selectedView.Value, selectedList.Value, itemUrl, Key, out _linkColumnDetails);
                        Session.Add(FormPrefix + "SharepointListData", sharepointListData);
                        Session.Add(FormPrefix + "SharepointItemListLinkColumnDetails", _linkColumnDetails);
                    }

                    if (sharepointListData.Rows.Count == 0)
                    {
                        displayModeErrorPanel.Visible = true;
                        displayModeErrorMessage.Text = "There are no items in the list.";
                    }
                    return sharepointListData;
                }
                else
                {
                    displayModeErrorPanel.Visible = true;
                    displayModeErrorTitile.Text = "Select Source";
                    displayModeErrorMessage.Text = "Please select the SharePoint information source you wish to access by using the Edit Content functionality located in the Component menu";
                    return sharepointListData;
                }
            }
            catch (System.Net.WebException webEx)
            {
                HttpWebResponse response = (HttpWebResponse)webEx.Response;
                if (response.StatusCode == HttpStatusCode.Unauthorized)
                {
                    displayModeErrorPanel.Visible = true;
                    displayModeErrorTitile.Text = "Invalid Credentials";
                    displayModeErrorMessage.Text = "You are not authorized to access this SharePoint content at this time. Please contact the Page Owner for more information.";
                }
                return sharepointListData;
            }
            catch (Exception ex)
            {
                displayModeErrorPanel.Visible = true;
                displayModeErrorMessage.Text = ex.Message;
                return sharepointListData;
            }
        }

        /// <summary>
        /// Bind the sharepoint items to the sharepointItemList Datagrid
        /// </summary>
        ///<param name="itemUrl">Server Url of the selected site or item</param>
        private void LoadSharepointItems(string itemUrl)
        {
            DataTable sharepointListData = RetrieveSharepointListData(itemUrl, false);
            if (sharepointListData.Rows.Count > 0)
            {
                sharepointItemList.DataSource = sharepointListData;
                sharepointItemList.DataBind();
            }
        }



        protected void sharepointItemListBound(object s, DataGridItemEventArgs e)
        {
            e.Item.Cells[1].Visible = false; //hide the itemId column in the datagrid
            _linkColumnDetails = (string)Session[FormPrefix + "SharepointItemListLinkColumnDetails"];

            if (e.Item.ItemType != ListItemType.Header & e.Item.ItemType != ListItemType.Footer)
            {
                if (!_linkColumnDetails.Equals(string.Empty))
                {
                    int linkPosition = int.Parse(_linkColumnDetails.Split('|')[1]); //position to re-add template column
                    string linkColumnName = _linkColumnDetails.Split('|')[0]; //column name to retrieve the link button text from the data table

                    //copy the current template column, remove it from the datagrid and re-add the template column in the required position
                    TableCell linkeColumn = e.Item.Cells[0];
                    e.Item.Cells.RemoveAt(linkPosition);
                    e.Item.Cells.AddAt(linkPosition, linkeColumn);
                    e.Item.Cells.RemoveAt(0);

                    //replace the text in the link button
                    LinkButton selectButton = (LinkButton)e.Item.FindControl("selectButton");
                    selectButton.Text = ((string)DataBinder.Eval(e.Item.DataItem, linkColumnName));
                }
                else
                {
                    e.Item.Cells.RemoveAt(0); //remove the link column if a link column does not exist
                }
            }
            else
            {
                e.Item.Cells.RemoveAt(0);
            }
        }


        protected void SelectButton_Click(object s, DataGridCommandEventArgs e)
        {
            if (e.CommandName.Equals("Select"))
            {
                CreateSharepointProxyObject();
                try
                {
                    string[] itemID = sharepointItemList.DataKeys[e.Item.ItemIndex].ToString().Split('|');
                    string contentType = itemID[0];
                    string itemUrl = itemID[1]; //server url of the selected item

                    //if the selected item is of type folder retrieve items within the folder
                    if (contentType.ToLower().Equals("folder"))
                    {
                        LoadSharepointItems(itemUrl);
                    }
                    //if the selected item is of type document download the document from the sharepoint site
                    else if (contentType.ToLower().Equals("document"))
                    {
                        string filename = itemUrl.Substring(itemUrl.IndexOf("/"));
                        string fileContentType = string.Empty;
                        byte[] fileStream = _Sharepoint.DownloadFile(filename, out fileContentType);

                        Response.Clear();
                        Response.AddHeader("Content-Disposition", "filename=" + filename);
                        Response.ContentType = fileContentType;

                        Response.BinaryWrite(fileStream);
                        Response.End();
                    }
                    //if the selected item is of type Task display Task details
                    else if (contentType.ToLower().Equals("task"))
                    {
                        TaskDetailPanel.Visible = true;
                        ItemDetailPanel.Visible = false;
                        string contentTypeId = itemID[2];
                        LoadTaskDetails(selectedList.Value, selectedView.Value, contentTypeId, itemUrl);
                    }
                }
                catch (System.Net.WebException webEx)
                {
                    HttpWebResponse response = (HttpWebResponse)webEx.Response;
                    if (response.StatusCode == HttpStatusCode.Unauthorized)
                    {
                        displayModeErrorPanel.Visible = true;
                        displayModeErrorTitile.Text = "Invalid Credentials";
                        displayModeErrorMessage.Text = "You are not authorized to access this SharePoint content at this time. Please contact the Page Owner for more information.";
                    }
                }
                catch (Exception ex)
                {
                    displayModeErrorPanel.Visible = true;
                    displayModeErrorMessage.Text = ex.Message;
                }
            }
        }


        /// <summary>
        /// Display details of a selected task in TaskDetailGrid
        /// </summary>
        /// <param name="listId">GUID of the selected list</param>
        /// <param name="viewId">GUID of the selected view</param>
        /// <param name="contentTypeId">ID of the selected items content type</param>
        /// <param name="itemUrl">Server Url of the selected item</param>
        private void LoadTaskDetails(string listId, string viewId, string contentTypeId, string itemUrl)
        {
            try
            {
                IEnumerable<KeyValuePair<string, string>> taskDetails = _Sharepoint.GetTaskDetails(listId, viewId, contentTypeId, itemUrl);
                TaskDetailGrid.DataSource = taskDetails;
                TaskDetailGrid.DataBind();
            }
            catch (Exception ex)
            {
                displayModeErrorPanel.Visible = true;
                displayModeErrorMessage.Text = ex.Message;
            }
        }


        /// <summary>
        /// Displays the Task Detail Panel which consists of details specific to a task
        /// </summary>
        protected void ItemDetailPanelButton_OnClick(object s, EventArgs e)
        {
            TaskDetailPanel.Visible = false;
            ItemDetailPanel.Visible = true;

            displayModeErrorPanel.Visible = false;

            DataTable sharepointListData = RetrieveSharepointListData(string.Empty, true);
            if (sharepointListData.Rows.Count > 0)
            {
                sharepointItemList.DataSource = sharepointListData;
                sharepointItemList.DataBind();
            }
        }


        /// <summary>
        /// Sort data in sharepointItemList Data Grid 
        /// </summary>
        protected void SharepointItemList_Sort(object s, DataGridSortCommandEventArgs e)
        {
            string sortDirection = string.Empty;

            if (Session[FormPrefix + "SortDirection"] != null)
            {
                sortDirection = ((string)Session[FormPrefix + "SortDirection"]);
                if (sortDirection.Equals("ASC"))
                {
                    sortDirection = "DESC";
                }
                else
                {
                    sortDirection = "ASC";
                }
            }


            DataTable sharepointListData = RetrieveSharepointListData(string.Empty, true);
            if (sharepointListData.Rows.Count > 0)
            {
                DataTable newDataTable = sharepointListData.Clone();

                if (sortDirection.Equals("ASC"))
                {
                    var tempEnumerableCollection = sharepointListData.AsEnumerable().OrderBy(dr => dr[e.SortExpression].ToString());
                    tempEnumerableCollection.CopyToDataTable(newDataTable, LoadOption.PreserveChanges);
                }
                else
                {
                    var tempEnumerableCollection = sharepointListData.AsEnumerable().OrderByDescending(dr => dr[e.SortExpression].ToString());
                    tempEnumerableCollection.CopyToDataTable(newDataTable, LoadOption.PreserveChanges);
                }


                sharepointItemList.DataSource = newDataTable;
                sharepointItemList.DataBind();

                //Session[FormPrefix + "SortExpression"] = e.SortExpression;
                Session[FormPrefix + "SortDirection"] = sortDirection;
            }
            
        }

#endregion

        
        private void ClearSelectedTreeViewValues()
        {
            selectedListTextBox.Text = string.Empty;
            selectedViewTextBox.Text = string.Empty;
            selectedSiteUrlTextBox.Text = string.Empty;
            selectedNodePathTextBox.Text = string.Empty;
        }

        private void PersistSelectedTreeViewValues()
        {
            selectedListTextBox.Text = selectedList.Value;
            selectedViewTextBox.Text = selectedView.Value;
            selectedSiteUrlTextBox.Text = selectedSiteUrl.Value;
            selectedNodePathTextBox.Text = selectedNode.Value;
        }

        /// <summary>
        /// Validate details entered in component properties
        /// </summary>
        /// <returns></returns>
        private bool ValidateComponentProperties(Mode mode)
        {
            try
            {
                string siteUrl = ValidSiteUrl(GetExternalComponentPropertyValue("SiteUrl"));
                NetworkCredential sharepointCredentials = new NetworkCredential(GetExternalComponentPropertyValue("Username"), GetExternalComponentPropertyValue("Password"), GetExternalComponentPropertyValue("Domain"));
                WebRequest request = WebRequest.Create(siteUrl);
                CredentialCache cache = new CredentialCache();
                cache.Add(new Uri(siteUrl), GetExternalComponentPropertyValue("AuthenticationType"), sharepointCredentials);
                request.Credentials = cache;

                WebResponse response = request.GetResponse();
                response.Close();
                return true;
            }
            catch (Exception ex)
            {
                DisplayErrorMessage(mode, ex);
                return false;
            }
        }


        private string ValidSiteUrl(string siteUrl)
        {
            if (siteUrl.ToLower().IndexOf("/default.aspx") > -1)
            {
                //remove default.aspx from the url and return the new url
                return siteUrl.Substring(0, siteUrl.ToLower().IndexOf("/default.aspx") + 1);
            }
            else if (siteUrl.ToLower().IndexOf("/allitems.aspx") > -1 && siteUrl.ToLower().IndexOf("/lists/") > -1)
            {
                return siteUrl.Substring(0, siteUrl.ToLower().IndexOf("/lists/") + 1);
            }
            else if (siteUrl.ToLower().IndexOf("/forms/allitems.aspx") > -1)
            {
                string url = siteUrl.Substring(0, siteUrl.ToLower().IndexOf("/forms/allitems.aspx"));
                return url.Substring(0, url.LastIndexOf("/") + 1);
            }
            return siteUrl;
        }

        private void DisplayErrorMessage(Mode mode, Exception ex)
        {
            switch (mode)
            {
                case Mode.Display:
                    displayModeErrorPanel.Visible = true;
                    displayModeErrorMessage.Text = ex.Message;
                    break;
                case Mode.Edit:
                    editModeErrorPanel.Visible = true;
                    editModeErrorMessage.Text = ex.Message;
                    break;
            }
        }


        private enum Mode
        {
            Display,
            Edit
        }
  
        
    }


    

}
