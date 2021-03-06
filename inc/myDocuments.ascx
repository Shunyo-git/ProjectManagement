<%@ Control Language="vb" AutoEventWireup="false" Codebehind="myDocuments.ascx.vb" Inherits="HsinYi.EIP.ProjectManagement.Web.myDocuments" TargetSchema="http://schemas.microsoft.com/intellisense/ie5" %>
<asp:datagrid id="dgDocs" AllowSorting="True" AutoGenerateColumns="False" CellPadding="2" Width="100%"
	BorderStyle="None" BorderColor="White" AllowPaging="True" runat="server" PageSize="5">
	<HeaderStyle Font-Bold="True" CssClass="grid-header"></HeaderStyle>
	<Columns>
		<asp:TemplateColumn SortExpression="Title" HeaderText="最新專案文件">
			<HeaderStyle HorizontalAlign="Left" Width="40%" CssClass="grid-header" VerticalAlign="Middle"></HeaderStyle>
			<ItemStyle CssClass="grid-first-item"></ItemStyle>
			<ItemTemplate>
				<a title="Document Detail" class="blacklineheight" href='DocumentDetail.aspx?DocumentID=<%#DataBinder.Eval(Container.DataItem, "DocumentID")%>&ProjectID=<%#DataBinder.Eval(Container.DataItem, "ProjectID")%>&index=-1&projectIndex=5'>
					<asp:label ID="Label2" Text='<%# DataBinder.Eval(Container, "DataItem.Title") %>' Runat="server" Visible="true" />
				</a>
			</ItemTemplate>
		</asp:TemplateColumn>
		<asp:TemplateColumn HeaderText="所屬專案">
			<HeaderStyle HorizontalAlign="Left" Width="25%" CssClass="grid-header" VerticalAlign="Middle"></HeaderStyle>
			<ItemStyle CssClass="grid-item"></ItemStyle>
			<ItemTemplate>
				<asp:label ID="Label4" Text='<%#  getProjectName(DataBinder.Eval(Container, "DataItem.ProjectID")) %>' Runat="server" Visible="true" />
			</ItemTemplate>
		</asp:TemplateColumn>
		<asp:TemplateColumn SortExpression="Category" HeaderText="文件類別">
			<HeaderStyle HorizontalAlign="Left" Width="15%" CssClass="grid-header" VerticalAlign="Middle"></HeaderStyle>
			<ItemStyle CssClass="grid-item"></ItemStyle>
			<ItemTemplate>
				<asp:label ID="Label1" Text='<%#  DataBinder.Eval(Container, "DataItem.CategoryName") %>' Runat="server" Visible="true" />
			</ItemTemplate>
		</asp:TemplateColumn>
		<asp:TemplateColumn SortExpression="CreatedUserID" HeaderText="新增者">
			<HeaderStyle HorizontalAlign="Left" Width="10%" CssClass="grid-header" VerticalAlign="Middle"></HeaderStyle>
			<ItemStyle CssClass="grid-last-item"></ItemStyle>
			<ItemTemplate>
				<asp:label ID="Label3" Text='<%# EIPSysSecurity.clsUser.getUserName(DataBinder.Eval(Container, "DataItem.CreatedUserID")) %>' Runat="server" Visible="true" />
			</ItemTemplate>
		</asp:TemplateColumn>
	</Columns>
	<PagerStyle HorizontalAlign="Right"></PagerStyle>
</asp:datagrid>
