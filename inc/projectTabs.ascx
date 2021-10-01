<%@ Control Language="vb" AutoEventWireup="false" Codebehind="projectTabs.ascx.vb" Inherits="HsinYi.EIP.ProjectManagement.Web.projectTabs" TargetSchema="http://schemas.microsoft.com/intellisense/ie5" %>
<TABLE id="Table2" cellSpacing="0" cellPadding="0" width="100%" border="0">
	<TR>
		<TD align="right">
			<asp:Label id="lblProjectName" runat="server" CssClass="item-text"></asp:Label>
		</TD>
	</TR>
</TABLE>
<TABLE id="Table1" cellSpacing="0" cellPadding="0" width="100%" border="0">
	<TR>
		<TD vAlign="bottom"><asp:DataList ID="dlProjectTabs" SelectedItemStyle-CssClass="admin-tab-active" ItemStyle-CssClass="admin-tab-inactive"
				CellSpacing="0" CellPadding="0" runat="server" EnableViewState="false" RepeatDirection="horizontal">
				<itemstyle Wrap="False"></itemstyle>
				<itemtemplate>
					<a href='<%= HsinYi.EIP.ProjectManagement.Web.Global.GetApplicationPath(Request) %>/<%# CType(Container.DataItem, Hsinyi.EIP.ProjectManagement.BusinessLogicLayer.TabItem).Path %>'>
						<%# CType(Container.DataItem, Hsinyi.EIP.ProjectManagement.BusinessLogicLayer.TabItem).Name %>
					</a>
				</itemtemplate>
				<selecteditemtemplate>
					<a  style="TEXT-DECORATION: none;" href='<%= HsinYi.EIP.ProjectManagement.Web.Global.GetApplicationPath(Request) %>/<%# CType(Container.DataItem, Hsinyi.EIP.ProjectManagement.BusinessLogicLayer.TabItem).Path %>'>
						<%# CType(Container.DataItem, Hsinyi.EIP.ProjectManagement.BusinessLogicLayer.TabItem).Name %>
					</a>
				</selecteditemtemplate>
			</asp:DataList></TD>
		<TD class="admin-tab-right" align="right" width="100%">&nbsp;&nbsp;</TD>
	</TR>
</TABLE>
