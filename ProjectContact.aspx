<%@ Register TagPrefix="uc1" TagName="projectTabs" Src="inc/projectTabs.ascx" %>
<%@ Register TagPrefix="uc1" TagName="myProjectTabs" Src="inc/mainTitleTabs.ascx" %>
<%@ Page Language="vb" AutoEventWireup="false" Codebehind="ProjectContact.aspx.vb" Inherits="HsinYi.EIP.ProjectManagement.Web.ProjectContact"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>專案連絡表</title>
		<meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
		<meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
		<meta name="vs_defaultClientScript" content="JavaScript">
		<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
		<LINK href="css/Styles.css" type="text/css" rel="stylesheet">
	</HEAD>
	<body>
		<FONT face="新細明體">
			<FORM id="ProjectList" method="post" runat="server">
				<TABLE id="Table1" cellSpacing="0" cellPadding="0" width="100%" border="0">
					<TR>
						<TD vAlign="bottom" bgColor="buttonface" height="46"><FONT face="新細明體">
								<uc1:myProjectTabs id="MyProjectTabs1" runat="server"></uc1:myProjectTabs></FONT></TD>
					</TR>
					<TR>
						<TD vAlign="top" height="15"><IMG height="15" src="images/spacer.gif" width="15"></TD>
					</TR>
				</TABLE>
				<TABLE id="Table2" cellSpacing="0" cellPadding="0" width="100%" border="0">
					<TR>
						<TD width="8"><IMG height="8" src="images/spacer.gif" width="8"></TD>
						<TD vAlign="top">
							<uc1:projectTabs id="ProjectTabs1" runat="server"></uc1:projectTabs>
							<TABLE class="admin-tan-border" id="Table3" height="530" cellSpacing="20" cellPadding="0"
								width="100%" border="0">
								<TR>
									<TD height="20">
										<P><FONT face="新細明體"></FONT>&nbsp;</P>
									</TD>
								</TR>
								<TR vAlign="top" height="*">
									<TD><FONT face="新細明體">
											<asp:datagrid id="Users" runat="server" BorderStyle="None" Width="100%" CellPadding="2" AutoGenerateColumns="False"
												BorderColor="White" AllowPaging="True">
												<HeaderStyle Font-Bold="True" CssClass="grid-header"></HeaderStyle>
												<Columns>
													
													
													<asp:TemplateColumn SortExpression="Name" HeaderText="組別">
														<HeaderStyle HorizontalAlign="Left" Width="15%" CssClass="grid-header" VerticalAlign="Middle"></HeaderStyle>
														<ItemStyle CssClass="grid-first-item"></ItemStyle>
														<ItemTemplate>
															 <%# DataBinder.Eval(Container, "DataItem.GroupName") %>
															 
														</ItemTemplate>
													</asp:TemplateColumn>
													
													
													<asp:TemplateColumn SortExpression="Name" HeaderText="姓名">
														<HeaderStyle HorizontalAlign="Left" Width="10%" CssClass="grid-header" VerticalAlign="Middle"></HeaderStyle>
														<ItemStyle CssClass="grid-item"></ItemStyle>
														<ItemTemplate>
															&nbsp;<a title="Edit User" class="blacklineheight" href='UserDetail.aspx?ProjectID=<%# DataBinder.Eval(Container, "DataItem.ProjectID") %>&MemberID=<%# DataBinder.Eval(Container, "DataItem.MemberID") %>&index=-1&projectIndex=2'><%# DataBinder.Eval(Container, "DataItem.DisplayName") %>
															</a>
														</ItemTemplate>
													</asp:TemplateColumn>
													<asp:TemplateColumn SortExpression="RoleName" HeaderText="角色">
														<HeaderStyle HorizontalAlign="Left" Width="15%" CssClass="grid-header" VerticalAlign="Middle"></HeaderStyle>
														<ItemStyle CssClass="grid-item"></ItemStyle>
														<ItemTemplate>
															&nbsp;
															<asp:label ID="lblGridRoleName" Text='<%# DataBinder.Eval(Container, "DataItem.RoleName") %>' Runat="server" Visible="true" />
														</ItemTemplate>
													</asp:TemplateColumn>
													<asp:TemplateColumn SortExpression="Department" HeaderText="單位">
														<HeaderStyle HorizontalAlign="Left" Width="15%" CssClass="grid-header" VerticalAlign="Middle"></HeaderStyle>
														<ItemStyle CssClass="grid-item"></ItemStyle>
														<ItemTemplate>
															&nbsp;
															<asp:label ID="Label1" Text='<%# DataBinder.Eval(Container, "DataItem.SourceUserDepartment") %>' Runat="server" Visible="true" />
														</ItemTemplate>
													</asp:TemplateColumn>
													<asp:TemplateColumn SortExpression="Title" HeaderText="分機">
														<HeaderStyle HorizontalAlign="Left" Width="10%" CssClass="grid-header" VerticalAlign="Middle"></HeaderStyle>
														<ItemStyle CssClass="grid-item"></ItemStyle>
														<ItemTemplate>
															&nbsp;
															<asp:label ID="Label2" Text='<%# DataBinder.Eval(Container, "DataItem.Extension") %>' Runat="server" Visible="true" />
														</ItemTemplate>
													</asp:TemplateColumn>
													<asp:TemplateColumn SortExpression="Title" HeaderText="樓層">
														<HeaderStyle HorizontalAlign="Left" Width="10%" CssClass="grid-header" VerticalAlign="Middle"></HeaderStyle>
														<ItemStyle CssClass="grid-item"></ItemStyle>
														<ItemTemplate>
															&nbsp;
															<asp:label ID="Label4" Text='<%# DataBinder.Eval(Container, "DataItem.Floor") %>' Runat="server" Visible="true" />
														</ItemTemplate>
													</asp:TemplateColumn>
													<asp:TemplateColumn SortExpression="Description" HeaderText="&lt;SPAN id=lblTel&gt;聯絡資料">
														<HeaderStyle HorizontalAlign="Left" Width="25%" CssClass="grid-header" VerticalAlign="Middle"></HeaderStyle>
														<ItemStyle CssClass="grid-item"></ItemStyle>
														<ItemTemplate>
															&nbsp;
															<asp:label ID="Label3" Text='<%# DataBinder.Eval(Container, "DataItem.Connects") %>' Runat="server" Visible="true" />
														</ItemTemplate>
													</asp:TemplateColumn>
													<asp:TemplateColumn Visible="False" HeaderText="UserID">
														<ItemTemplate>
															&nbsp;
															<asp:label ID="lblGridUserID" Text='<%# DataBinder.Eval(Container, "DataItem.UserID") %>' Runat="server" Visible="true" />
														</ItemTemplate>
													</asp:TemplateColumn>
													<asp:TemplateColumn Visible="False" HeaderText="UserID">
														<ItemTemplate>
															&nbsp;
															<asp:label ID="lblGridUserName" Text='<%# DataBinder.Eval(Container, "DataItem.UserName")  %>' Runat="server" Visible="true" />
														</ItemTemplate>
													</asp:TemplateColumn>
												</Columns>
												<PagerStyle HorizontalAlign="Center"></PagerStyle>
											</asp:datagrid></FONT></TD>
								</TR>
								<TR vAlign="top">
									<TD colSpan="3" height="12"></TD>
								</TR>
							</TABLE>
						</TD>
						<TD width="11"><IMG height="11" src="images/spacer.gif" width="11"></TD>
					</TR>
					<TR>
						<TD vAlign="top" colSpan="5" height="15"><IMG height="15" src="images/spacer.gif" width="15"></TD>
					</TR>
				</TABLE>
			</FORM>
		</FONT>
	</body>
</HTML>
