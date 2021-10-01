<%@ Register TagPrefix="uc1" TagName="myProjectTabs" Src="inc/mainTitleTabs.ascx" %>
<%@ Register TagPrefix="uc1" TagName="projectTabs" Src="inc/projectTabs.ascx" %>
<%@ Page Language="vb" AutoEventWireup="false" Codebehind="ProjectList.aspx.vb" Inherits="HsinYi.EIP.ProjectManagement.Web.ProjectList"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>�M�צC��</title>
		<meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
		<meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
		<meta content="JavaScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
		<LINK href="css/Styles.css" type="text/css" rel="stylesheet">
	</HEAD>
	<body>
		<FONT face="�s�ө���">
			<FORM id="ProjectList" method="post" runat="server">
				<table cellSpacing="0" cellPadding="0" width="100%" border="0">
					<tr>
						<td vAlign="bottom" bgColor="buttonface" height="46"><FONT face="�s�ө���"><uc1:myprojecttabs id="MyProjectTabs1" runat="server"></uc1:myprojecttabs></FONT></td>
					</tr>
					<tr>
						<td vAlign="top" height="15"><IMG height="15" src="images/spacer.gif" width="15"></td>
					</tr>
				</table>
				<TABLE id="Table2" cellSpacing="0" cellPadding="0" width="100%" border="0">
					<TR>
						<TD width="8"><IMG height="8" src="images/spacer.gif" width="8"></TD>
						<TD vAlign="top">
							<TABLE class="admin-tan-border" id="Table3" height="530" cellSpacing="20" cellPadding="0"
								width="100%" border="0">
								<TR>
									<TD height="20"><asp:button id="NewProjectButton" runat="server" Text="Create New Project" CssClass="standard-text"
											Visible="False"></asp:button></TD>
								</TR>
								<TR vAlign="top" height="*">
									<TD><asp:datagrid id="ProjectsGrid" runat="server" AllowSorting="True" AutoGenerateColumns="False"
											CellPadding="2" Width="100%" BorderStyle="None" BorderColor="White" allowpaging="True" OnPageIndexChanged="ProjectsGrid_Page">
											<HeaderStyle CssClass="grid-header"></HeaderStyle>
											<Columns>
												<asp:HyperLinkColumn DataNavigateUrlField="ProjectID" DataNavigateUrlFormatString="ProjectDetails.aspx?projectID={0}&amp;index=-1&amp;projectIndex=0"
													DataTextField="ProjectName" SortExpression="ProjectName" HeaderText="�M�צW��">
													<HeaderStyle HorizontalAlign="Left" Width="40%" CssClass="grid-header"></HeaderStyle>
													<ItemStyle CssClass="grid-first-item"></ItemStyle>
												</asp:HyperLinkColumn>
												<asp:TemplateColumn SortExpression="Manager" HeaderText="�t�d�H">
													<HeaderStyle HorizontalAlign="Left" Width="12%" CssClass="grid-header"></HeaderStyle>
													<ItemStyle CssClass="grid-item"></ItemStyle>
													<ItemTemplate>
														&nbsp;<%# EIPSysSecurity.clsUser.getUserName(DataBinder.Eval(Container, "DataItem.ManagerUserID")) %>
													</ItemTemplate>
												</asp:TemplateColumn>
												<asp:TemplateColumn SortExpression="ProjectChief" HeaderText="�D�n�����">
													<HeaderStyle HorizontalAlign="Left" Width="12%" CssClass="grid-header"></HeaderStyle>
													<ItemStyle HorizontalAlign="Left" CssClass="grid-item"></ItemStyle>
													<ItemTemplate>
														&nbsp;<%# EIPSysSecurity.clsUser.getUserName(DataBinder.Eval(Container, "DataItem.ChiefUserID" )) %>
													</ItemTemplate>
												</asp:TemplateColumn>
												<asp:TemplateColumn SortExpression="StartDate" HeaderText="�}�l�ɶ�">
													<HeaderStyle HorizontalAlign="Left" Width="12%" CssClass="grid-header"></HeaderStyle>
													<ItemStyle HorizontalAlign="Left" CssClass="grid-item"></ItemStyle>
													<ItemTemplate>
														&nbsp;<%# Date.Parse( DataBinder.Eval(Container, "DataItem.StartDate")).ToShortDateString() %>
													</ItemTemplate>
												</asp:TemplateColumn>
												<asp:TemplateColumn SortExpression="EndDate" HeaderText="�����ɶ�">
													<HeaderStyle HorizontalAlign="Left" Width="12%" CssClass="grid-header"></HeaderStyle>
													<ItemStyle HorizontalAlign="Left" CssClass="grid-item"></ItemStyle>
													<ItemTemplate>
														&nbsp;<%# Date.Parse(DataBinder.Eval(Container, "DataItem.EndDate")).ToShortDateString() %>
													</ItemTemplate>
												</asp:TemplateColumn>
												<asp:TemplateColumn SortExpression="State" HeaderText="���A">
													<HeaderStyle HorizontalAlign="Right" Width="12%" CssClass="grid-header"></HeaderStyle>
													<ItemStyle HorizontalAlign="Right" CssClass="grid-last-item"></ItemStyle>
													<ItemTemplate>
														&nbsp
														<%#  getStateColorName(DataBinder.Eval(Container, "DataItem.StateID"),DataBinder.Eval(Container, "DataItem.StateName") )%>
													</ItemTemplate>
												</asp:TemplateColumn>
											</Columns>
											<PagerStyle Visible="False" HorizontalAlign="Center" PageButtonCount="5" Mode="NumericPages"></PagerStyle>
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
		</FONT>
	</body>
</HTML>