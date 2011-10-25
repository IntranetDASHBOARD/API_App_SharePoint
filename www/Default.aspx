<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Default.aspx.cs" Inherits="SharePointConnector.Default" EnableEventValidation="False"%>
<%@ Register Assembly="System.Web.Extensions, Version=3.5.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35"
    Namespace="System.Web.UI.WebControls" TagPrefix="asp" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>SharePoint Connector</title>
    <% Response.Write("<link href=\"" + Intranet.CurrentSubsiteThemePath + "\" type=\"text/css\" rel=\"stylesheet\">"); %>
    <link type="text/css" rel="Stylesheet" href="includes/SharePointConnectorStyles.css" />
    <link type="text/css" rel="Stylesheet" href="/includes/css/Core.aspx" />
	<link href="/includes/css/api_styleReset.aspx" type="text/css" rel="stylesheet" />
	<script src="/includes/js/jquery.aspx" type="text/javascript"></script>
</head>
<body class="noBodyBackground">
    <form id="form1" runat="server">
    <!-- Display View -->
		<asp:PlaceHolder ID="DisplayMode" runat="server">
		    <asp:Panel ID="ItemDetailPanel" runat="server" Visible="true">
                <asp:DataGrid ID="sharepointItemList" CssClass="tableMain sharepointDatagrid bCollapseSeparate" runat="server" CellPadding="3" CellSpacing="0" Width="100%" DataKeyField="ItemID" OnItemDataBound="sharepointItemListBound" AutoGenerateColumns="true" HeaderStyle-HorizontalAlign="Left" OnItemCommand="SelectButton_Click" AllowSorting="true" OnSortCommand="SharepointItemList_Sort" UseAccessibleHeader="true">
                    <Columns>
                        <asp:TemplateColumn>
                            <ItemTemplate>
                                <asp:LinkButton ID="selectButton" runat="server" Text="Select" CommandName="Select" />
                            </ItemTemplate>
                        </asp:TemplateColumn>
                    </Columns>
                </asp:DataGrid>    
	        </asp:Panel>
	        <asp:Panel ID="TaskDetailPanel" runat="server" Visible="false">
	            <table>
	                <tr>
	                    <td>
	                    <asp:DataGrid runat="server" ID="TaskDetailGrid" CssClass="tableMain bCollapseSeparate" Width="100%" CellPadding="3" AutoGenerateColumns="true" ShowHeader="false" UseAccessibleHeader="true"></asp:DataGrid>
	                    </td>
	                </tr>
	                <tr>
	                    <td>
	                    <br />
	                    <asp:Button ID="ItemDetailPanelButton" runat="server" CssClass="button" Text="Back to List" OnClick="ItemDetailPanelButton_OnClick" />
	                    </td>
	                </tr>
	            </table>
	        </asp:Panel>
	        <asp:Panel runat="server" id="displayModeErrorPanel" CssClass="errorLabel" visible="false">
                <asp:Label ID="displayModeErrorTitile" CssClass="errorTitle" runat="server"></asp:Label>
                <asp:Label ID="displayModeErrorMessage" CssClass="errorSubTitle" runat="server"></asp:Label>
            </asp:Panel>
		</asp:PlaceHolder>
	
	<!-- Edit View -->
		<asp:PlaceHolder ID="EditMode" runat="server">
            <asp:TreeView ID="SiteDetailTreeView" runat="server" ExpandDepth="0" Font-Size="11px" ShowLines="true" OnTreeNodeExpanded="SiteDetailTreeView_Expanded" PopulateNodesFromClient="true" OnSelectedNodeChanged="SiteDetailTreeView_NodeSelected" PathSeparator="!">
                <RootNodeStyle Font-Bold="true" />
                <ParentNodeStyle Font-Bold="false" />
                <SelectedNodeStyle Font-Bold="true" />
            </asp:TreeView>
            
            <asp:Panel runat="server" ID="editModeErrorPanel" CssClass="errorLabel" Visible="false">
                <asp:Label ID="editModeErrorTitle" CssClass="errorTitle" runat="server" Text="Error"></asp:Label>
                <asp:Label ID="editModeErrorMessage" CssClass="errorSubTitle" runat="server"></asp:Label>
            </asp:Panel>
            
            <asp:TextBox ID="selectedListTextBox" runat="server" style="display:none"></asp:TextBox>
            <asp:TextBox ID="selectedViewTextBox" runat="server" style="display:none"></asp:TextBox>
            <asp:TextBox ID="selectedSiteUrlTextBox" runat="server" style="display:none"></asp:TextBox>
            <asp:TextBox ID="selectedNodePathTextBox" runat="server" style="display:none"></asp:TextBox>
            <asp:TextBox ID="propertyDetailsTextBox" runat="server" style="display:none"></asp:TextBox>
        
            <asp:Literal ID="scriptLiteral" runat="server" ></asp:Literal>    
		</asp:PlaceHolder>
		
		
	</form>
</body>
<script type="text/javascript">
    $(document).ready(function() {

        $(".sharepointDatagrid th, #TaskDetailGrid td:nth-child(1)").addClass("tableHeader");

        $(".sharepointDatagrid td, #TaskDetailGrid td:nth-child(2)").addClass("tableItem");

        $(".sharepointDatagrid tr:last-child td, #TaskDetailGrid tr:last-child td").addClass("noBottomBorder");

    });

    function radioButtonClicked(item) {
        document.getElementById(item.parentNode.id).click();
    }
</script>
</html>            
