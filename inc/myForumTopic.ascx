<%@ Control Language="vb" AutoEventWireup="false" Codebehind="myForumTopic.ascx.vb" Inherits="HsinYi.EIP.ProjectManagement.Web.myForumTopic" TargetSchema="http://schemas.microsoft.com/intellisense/ie5" %>
<asp:DataGrid ID="dgList" AllowPaging="True" BorderColor="White" AllowSorting="True" AutoGenerateColumns="False"
	CellPadding="2" Width="100%" BorderStyle="None" runat="server" PageSize="5">
  <headerstyle Font-Bold="True" CssClass="grid-header"></headerstyle>
  <columns>
  <asp:TemplateColumn SortExpression="Subject" HeaderText="最新討論話題">
    <headerstyle HorizontalAlign="Left" Width="50%" CssClass="grid-header" VerticalAlign="Middle"></headerstyle>
    <itemstyle CssClass="grid-first-item"></itemstyle>
    <itemtemplate> <a title="Event Detail" class="blacklineheight" href='ForumTopicDetail.aspx?TopicID=<%#DataBinder.Eval(Container.DataItem, "TopicID")%>&ProjectID=<%#DataBinder.Eval(Container.DataItem, "ProjectID")%>&index=-1&projectIndex=6'>
      <asp:Label ID="Label2" Text='<%# DataBinder.Eval(Container, "DataItem.Subject") %>' runat="server" Visible="true" />
    </a> </itemtemplate>
  </asp:TemplateColumn>
  <asp:TemplateColumn HeaderText="所屬專案">
    <headerstyle HorizontalAlign="Left" Width="30%" CssClass="grid-header" VerticalAlign="Middle"></headerstyle>
    <itemstyle CssClass="grid-item"></itemstyle>
    <itemtemplate>
      <asp:Label ID="Label3" Text='<%# getProjectName(DataBinder.Eval(Container, "DataItem.ProjectID"))  %>' runat="server" Visible="true" />
    </itemtemplate>
  </asp:TemplateColumn>
  <asp:TemplateColumn SortExpression="LastPostDate" HeaderText="最新回應">
    <headerstyle HorizontalAlign="Right" Width="20%" CssClass="grid-header" VerticalAlign="Middle"></headerstyle>
    <itemstyle Wrap="False" HorizontalAlign="Right" CssClass="grid-last-item"></itemstyle>
    <itemtemplate> <%# DataBinder.Eval(Container, "DataItem.LastPostDate") %> <br>
      By <%# EIPSysSecurity.clsUser.getUserName( DataBinder.Eval(Container, "DataItem.LastPostUserID") )%> </itemtemplate>
  </asp:TemplateColumn>
  </columns>
  <pagerstyle HorizontalAlign="Right"></pagerstyle>
</asp:DataGrid>
