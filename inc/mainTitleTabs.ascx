<%@ Control Language="vb" AutoEventWireup="false" Codebehind="mainTitleTabs.ascx.vb" Inherits="HsinYi.EIP.ProjectManagement.Web.mainTitleTabs" TargetSchema="http://schemas.microsoft.com/intellisense/ie5" %>
<table cellSpacing="0" cellPadding="0" width="100%" border="0">
	<tr>
		<td vAlign="bottom"><asp:datalist id="tabs" SelectedItemStyle-CssClass="tab-active" ItemStyle-CssClass="tab-inactive"
				CellSpacing="0" CellPadding="0" runat="server" EnableViewState="false" RepeatDirection="horizontal">
				<ItemStyle Wrap="False"></ItemStyle>
				<itemtemplate>
					<a href='<%= HsinYi.EIP.ProjectManagement.Web.Global.GetApplicationPath(Request) %>/<%# CType(Container.DataItem, Hsinyi.EIP.ProjectManagement.BusinessLogicLayer.TabItem).Path %>'>
						<%# CType(Container.DataItem, Hsinyi.EIP.ProjectManagement.BusinessLogicLayer.TabItem).Name %>
					</a>
				</itemtemplate>
				<selecteditemtemplate>
					<%# CType(Container.DataItem, Hsinyi.EIP.ProjectManagement.BusinessLogicLayer.TabItem).Name %>
				</selecteditemtemplate>
			</asp:datalist></td>
		<TD align="right" width="100%">µn¤JªÌ¡G
			<asp:label id="lblCurrentUser" runat="server"></asp:label>&nbsp;</TD>
	</tr>
</table>
