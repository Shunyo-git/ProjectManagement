<%@ Control Language="vb" AutoEventWireup="false" Codebehind="DocCategoriesTabs.ascx.vb" Inherits="HsinYi.EIP.ProjectManagement.Web.DocCategoriesTabs" TargetSchema="http://schemas.microsoft.com/intellisense/ie5" %>
<TABLE id="Table1" cellSpacing="0" cellPadding="0" width="100%" border="0">
	<TR>
		<TD vAlign="bottom">
			<asp:datalist id="dlProjectTabs" SelectedItemStyle-CssClass="admin-tab-active" ItemStyle-CssClass="admin-tab-inactive"
				CellSpacing="0" CellPadding="0" runat="server" EnableViewState="false" RepeatDirection="horizontal">
				<SelectedItemStyle CssClass="admin-tab-active"></SelectedItemStyle>
				<SelectedItemTemplate>
					<A style="TEXT-DECORATION: none" href="<%= HsinYi.EIP.ProjectManagement.Web.Global.GetApplicationPath(Request) %>/<%# CType(Container.DataItem, Hsinyi.EIP.ProjectManagement.BusinessLogicLayer.TabItem).Path %>">
						<%# CType(Container.DataItem, Hsinyi.EIP.ProjectManagement.BusinessLogicLayer.TabItem).Name %>
					</A>
				</SelectedItemTemplate>
				<ItemStyle Wrap="False" CssClass="admin-tab-inactive"></ItemStyle>
				<ItemTemplate>
					<A style="TEXT-DECORATION: none" href="<%= HsinYi.EIP.ProjectManagement.Web.Global.GetApplicationPath(Request) %>/<%# CType(Container.DataItem, Hsinyi.EIP.ProjectManagement.BusinessLogicLayer.TabItem).Path %>">
						<%# CType(Container.DataItem, Hsinyi.EIP.ProjectManagement.BusinessLogicLayer.TabItem).Name %>
					</A>
				</ItemTemplate>
			</asp:datalist></TD>
		<TD class="admin-tab-right" align="right" width="100%">&nbsp;&nbsp;</TD>
	</TR>
</TABLE>
