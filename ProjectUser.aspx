<%@ Register TagPrefix="uc1" TagName="myProjectTabs" Src="inc/mainTitleTabs.ascx" %>
<%@ Register TagPrefix="uc1" TagName="projectTabs" Src="inc/projectTabs.ascx" %>
<%@ Page Language="vb" AutoEventWireup="false" Codebehind="ProjectUser.aspx.vb" Inherits="HsinYi.EIP.ProjectManagement.Web.ProjectUser"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>�M�צ���</title>
		<meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
		<meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
		<meta content="JavaScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
		<LINK href="css/Styles.css" type="text/css" rel="stylesheet">
	</HEAD>
	<body>
		<FORM id="ProjectList" method="post" runat="server">
			<TABLE id="Table1" cellSpacing="0" cellPadding="0" width="100%" border="0">
				<TR>
					<TD vAlign="bottom" bgColor="buttonface" height="46"><FONT face="�s�ө���"><uc1:myprojecttabs id="MyProjectTabs2" runat="server"></uc1:myprojecttabs></FONT></TD>
				</TR>
				<TR>
					<TD vAlign="top" height="15"><IMG height="15" src="images/spacer.gif" width="15"></TD>
				</TR>
			</TABLE>
			<TABLE id="Table2" cellSpacing="0" cellPadding="0" width="100%" border="0">
				<TR>
					<TD width="8"><IMG height="8" src="images/spacer.gif" width="8"></TD>
					<TD vAlign="top"><uc1:projecttabs id="ProjectTabs2" runat="server"></uc1:projecttabs>
						<TABLE class="admin-tan-border" id="Table3" height="530" cellSpacing="20" cellPadding="0"
							width="100%" border="0">
							<TR>
								<TD height="20">
									<asp:panel id="PanelModify" runat="server">
										<TABLE id="Table6" cellSpacing="0" cellPadding="0" width="100%" border="0">
											<TR>
												<TD style="HEIGHT: 14px"></TD>
												<TD style="HEIGHT: 14px" noWrap align="right"><IMG alt="" src="images/icon01.gif" align="absBottom">&nbsp;
													<asp:LinkButton id="lbtnNew" runat="server">�s�W����</asp:LinkButton>&nbsp;
												</TD>
											</TR>
											<TR>
												<TD noWrap align="right" colSpan="2"><IMG alt="" src="images/l_ba_bottom.GIF">
												</TD>
											</TR>
										</TABLE>
									</asp:panel></TD>
							</TR>
							<TR vAlign="top" height="*">
								<TD><asp:datagrid id="Users" runat="server" BorderStyle="None" Width="100%" CellPadding="2" AutoGenerateColumns="False"
										AllowSorting="True" BorderColor="White" AllowPaging="True">
										<HeaderStyle Font-Bold="True" CssClass="grid-header"></HeaderStyle>
										<Columns>
											<asp:TemplateColumn SortExpression="Name" HeaderText="�m�W">
												<HeaderStyle HorizontalAlign="Left" Width="15%" CssClass="grid-header" VerticalAlign="Middle"></HeaderStyle>
												<ItemStyle CssClass="grid-first-item"></ItemStyle>
												<ItemTemplate>
													&nbsp;<a title="Edit User" class="blacklineheight" href='UserDetail.aspx?ProjectID=<%# DataBinder.Eval(Container, "DataItem.ProjectID") %>&MemberID=<%# DataBinder.Eval(Container, "DataItem.MemberID") %>&index=-1&projectIndex=2'><%# DataBinder.Eval(Container, "DataItem.DisplayName") %>
													</a>
												</ItemTemplate>
											</asp:TemplateColumn>
											<asp:TemplateColumn SortExpression="RoleName" HeaderText="����">
												<HeaderStyle HorizontalAlign="Left" Width="15%" CssClass="grid-header" VerticalAlign="Middle"></HeaderStyle>
												<ItemStyle CssClass="grid-item"></ItemStyle>
												<ItemTemplate>
													&nbsp;
													<asp:label ID="lblGridRoleName" Text='<%# DataBinder.Eval(Container, "DataItem.RoleName") %>' Runat="server" Visible="true" />
												</ItemTemplate>
												<EditItemTemplate>
													<asp:DropDownList Width="100px" ID="ddlRoles" CssClass="Standard-text" DataSource='<%# GetRoles() %>' DataTextField="Name" DataValueField="RoleID" Runat="server" />
												</EditItemTemplate>
											</asp:TemplateColumn>
											<asp:TemplateColumn SortExpression="Department" HeaderText="���ݳ��">
												<HeaderStyle HorizontalAlign="Left" Width="15%" CssClass="grid-header" VerticalAlign="Middle"></HeaderStyle>
												<ItemStyle CssClass="grid-item"></ItemStyle>
												<ItemTemplate>
													&nbsp;
													<asp:label ID="Label1" Text='<%# DataBinder.Eval(Container, "DataItem.SourceUserDepartment") %>' Runat="server" Visible="true" />
												</ItemTemplate>
												<EditItemTemplate>
												</EditItemTemplate>
											</asp:TemplateColumn>
											<asp:TemplateColumn SortExpression="Title" HeaderText="¾��">
												<HeaderStyle HorizontalAlign="Left" Width="15%" CssClass="grid-header" VerticalAlign="Middle"></HeaderStyle>
												<ItemStyle CssClass="grid-item"></ItemStyle>
												<ItemTemplate>
													&nbsp;
													<asp:label ID="Label2" Text='<%# DataBinder.Eval(Container, "DataItem.SourceUserTitle") %>' Runat="server" Visible="true" />
												</ItemTemplate>
												<EditItemTemplate>
												</EditItemTemplate>
											</asp:TemplateColumn>
											<asp:TemplateColumn SortExpression="Description" HeaderText="����">
												<HeaderStyle HorizontalAlign="Left" Width="40%" CssClass="grid-header" VerticalAlign="Middle"></HeaderStyle>
												<ItemStyle CssClass="grid-item"></ItemStyle>
												<ItemTemplate>
													&nbsp;
													<asp:label ID="Label3" Text='<%# DataBinder.Eval(Container, "DataItem.Description") %>' Runat="server" Visible="true" />
												</ItemTemplate>
												<EditItemTemplate>
												</EditItemTemplate>
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
										<PagerStyle Visible="False" HorizontalAlign="Center"></PagerStyle>
									</asp:datagrid><BR>
									<TABLE id="Table4" cellSpacing="0" cellPadding="0" width="100%" border="0">
										<TR>
											<TD style="FONT-SIZE: 11px" align="center">
												<asp:linkbutton id="lbFirstPage" onclick="PagerButtonClick" runat="server" CommandArgument="�̫e��">�̫e��</asp:linkbutton><FONT face="�s�ө���">&nbsp;&nbsp;
												</FONT>
												<asp:linkbutton id="lbLastPage" onclick="PagerButtonClick" runat="server" CommandArgument="�̥���">�̥���</asp:linkbutton><FONT face="�s�ө���">&nbsp;&nbsp; 
													&lt;&lt;&nbsp;&nbsp; �� </FONT>
												<asp:label id="lblNowPage" runat="server"></asp:label><FONT face="�s�ө���">&nbsp;�� , �@ </FONT>
												<asp:label id="lblTotalPage" runat="server"></asp:label><FONT face="�s�ө���">&nbsp;��&nbsp;&nbsp;&nbsp;</FONT><FONT face="�s�ө���">
													&gt;&gt;&nbsp;&nbsp;</FONT>
												<asp:linkbutton id="lbPrevPage" onclick="PagerButtonClick" runat="server" CommandArgument="�W�@��">�W�@��</asp:linkbutton><FONT face="�s�ө���">&nbsp;&nbsp;
												</FONT>
												<asp:linkbutton id="lbNextPage" onclick="PagerButtonClick" runat="server" CommandArgument="�U�@��">�U�@��</asp:linkbutton><FONT face="�s�ө���"></FONT></TD>
											<TD style="FONT-SIZE: 11px" align="right"><FONT face="�s�ө���">&nbsp;�`�p
													<asp:Label id="lblRecordCount" runat="server"></asp:Label><FONT face="�s�ө���">&nbsp;��&nbsp;
													</FONT></FONT>
											</TD>
										</TR>
									</TABLE>
								</TD>
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
	</body>
</HTML>
