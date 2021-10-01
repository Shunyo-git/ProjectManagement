<%@ Control Language="vb" AutoEventWireup="false" Codebehind="myProjectList.ascx.vb" Inherits="HsinYi.EIP.ProjectManagement.Web.myProjectList" TargetSchema="http://schemas.microsoft.com/intellisense/ie5" %>
<asp:datagrid id="ProjectsGrid" AutoGenerateColumns="False" CellPadding="2" Width="100%" BorderStyle="None"
	BorderColor="White" AllowPaging="True" PageSize="5" runat="server">
	<HeaderStyle CssClass="grid-header"></HeaderStyle>
	<Columns>
		<asp:HyperLinkColumn DataNavigateUrlField="ProjectID" DataNavigateUrlFormatString="../ProjectDetails.aspx?projectID={0}&amp;index=-1&amp;projectIndex=0"
			DataTextField="ProjectName" SortExpression="ProjectName" HeaderText="我的專案列表 ">
			<HeaderStyle HorizontalAlign="Left" Width="50%" CssClass="grid-header"></HeaderStyle>
			<ItemStyle CssClass="grid-first-item"></ItemStyle>
		</asp:HyperLinkColumn>
		<asp:TemplateColumn SortExpression="Manager" HeaderText="負責人">
			<HeaderStyle HorizontalAlign="Left" Width="15%" CssClass="grid-header"></HeaderStyle>
			<ItemStyle CssClass="grid-item"></ItemStyle>
			<ItemTemplate>
				&nbsp;<%# EIPSysSecurity.clsUser.getUserName(DataBinder.Eval(Container, "DataItem.ManagerUserID")) %>
			</ItemTemplate>
		</asp:TemplateColumn>
		<asp:TemplateColumn SortExpression="ProjectChief" HeaderText="主要執行者">
			<HeaderStyle HorizontalAlign="Left" Width="20%" CssClass="grid-header"></HeaderStyle>
			<ItemStyle HorizontalAlign="Left" CssClass="grid-item"></ItemStyle>
			<ItemTemplate>
				&nbsp;<%# EIPSysSecurity.clsUser.getUserName(DataBinder.Eval(Container, "DataItem.ChiefUserID" )) %>
			</ItemTemplate>
		</asp:TemplateColumn>
		<asp:TemplateColumn SortExpression="State" HeaderText="狀態">
			<HeaderStyle HorizontalAlign="Right" Width="15%" CssClass="grid-header"></HeaderStyle>
			<ItemStyle HorizontalAlign="Right" CssClass="grid-last-item"></ItemStyle>
			<ItemTemplate>
				&nbsp
				<%#  getStateColorName(DataBinder.Eval(Container, "DataItem.StateID"),DataBinder.Eval(Container, "DataItem.StateName") )%>
			</ItemTemplate>
		</asp:TemplateColumn>
	</Columns>
	<PagerStyle HorizontalAlign="Right" PageButtonCount="5"></PagerStyle>
</asp:datagrid>
