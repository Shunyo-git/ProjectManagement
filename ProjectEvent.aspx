<%@ Page Language="vb" AutoEventWireup="false" Codebehind="ProjectEvent.aspx.vb" Inherits="HsinYi.EIP.ProjectManagement.Web.ProjectEvent"%>
<%@ Register TagPrefix="uc1" TagName="projectTabs" Src="inc/projectTabs.ascx" %>
<%@ Register TagPrefix="uc1" TagName="myProjectTabs" Src="inc/mainTitleTabs.ascx" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>特殊事件</title>
		<meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
		<meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
		<meta content="JavaScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
		<LINK href="css/Styles.css" type="text/css" rel="stylesheet">
	</HEAD>
	<body>
		<FONT face="新細明體">
			<FORM id="EventList" method="post" runat="server">
				<TABLE id="Table1" cellSpacing="0" cellPadding="0" width="100%" border="0">
					<TR>
						<TD vAlign="bottom" bgColor="buttonface" height="46"><FONT face="新細明體"><uc1:myprojecttabs id="MyProjectTabs1" runat="server"></uc1:myprojecttabs></FONT></TD>
					</TR>
					<TR>
						<TD vAlign="top" height="15"><IMG height="15" src="images/spacer.gif" width="15"></TD>
					</TR>
				</TABLE>
				<TABLE id="Table2" cellSpacing="0" cellPadding="0" width="100%" border="0">
					<TR>
						<TD width="8"><IMG height="8" src="images/spacer.gif" width="8"></TD>
						<TD vAlign="top"><uc1:projecttabs id="ProjectTabs1" runat="server"></uc1:projecttabs>
							<TABLE class="admin-tan-border" id="Table3" height="530" cellSpacing="20" cellPadding="0"
								width="100%" border="0">
								<TR>
									<TD height="20"><asp:panel id="PanelModify" runat="server">
											<TABLE id="Table6" cellSpacing="0" cellPadding="0" width="100%" border="0">
												<TR>
													<TD style="HEIGHT: 14px"><FONT face="新細明體"></FONT></TD>
													<TD style="HEIGHT: 14px" noWrap align="right"><IMG alt="" src="images/icon01.gif" align="absBottom">&nbsp;
														<asp:LinkButton id="lbtnNew" runat="server">新增記錄</asp:LinkButton>&nbsp;
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
									<TD><FONT face="新細明體"><asp:datagrid id="dgEvents" runat="server" AllowPaging="True" BorderColor="White" AllowSorting="True"
												AutoGenerateColumns="False" CellPadding="2" Width="100%" BorderStyle="None">
												<HeaderStyle Font-Bold="True" CssClass="grid-header"></HeaderStyle>
												<Columns>
													<asp:TemplateColumn SortExpression="EventDate" HeaderText="記錄時間">
														<HeaderStyle HorizontalAlign="Left" Width="10%" CssClass="grid-header" VerticalAlign="Middle"></HeaderStyle>
														<ItemStyle CssClass="grid-first-item"></ItemStyle>
														<ItemTemplate>
															<asp:label ID="Label1" Text='<%# Date.Parse( DataBinder.Eval(Container, "DataItem.EventDate")).ToShortDateString %>' Runat="server" Visible="true" />
														</ItemTemplate>
													</asp:TemplateColumn>
													<asp:TemplateColumn SortExpression="Topic" HeaderText="記錄主題">
														<HeaderStyle HorizontalAlign="Left" Width="60%" CssClass="grid-header" VerticalAlign="Middle"></HeaderStyle>
														<ItemStyle CssClass="grid-item"></ItemStyle>
														<ItemTemplate>
															<a title="Event Detail" class="blacklineheight" href='EventDetail.aspx?EventID=<%#DataBinder.Eval(Container.DataItem, "EventID")%>&ProjectID=<%#DataBinder.Eval(Container.DataItem, "ProjectID")%>&index=-1&projectIndex=4'>
																<asp:label ID="Label2" Text='<%# DataBinder.Eval(Container, "DataItem.Topic") %>' Runat="server" Visible="true" />
															</a>
														</ItemTemplate>
													</asp:TemplateColumn>
													<asp:TemplateColumn SortExpression="CreatedUserID" HeaderText="作者">
														<HeaderStyle HorizontalAlign="Left" Width="10%" CssClass="grid-header" VerticalAlign="Middle"></HeaderStyle>
														<ItemStyle CssClass="grid-item"></ItemStyle>
														<ItemTemplate>
															<asp:label ID="Label3" Text='<%# EIPSysSecurity.clsUser.getUserName(DataBinder.Eval(Container, "DataItem.CreatedUserID")) %>' Runat="server" Visible="true" />
														</ItemTemplate>
													</asp:TemplateColumn>
													<asp:TemplateColumn SortExpression="CreatedDate" HeaderText="新增時間">
														<HeaderStyle HorizontalAlign="Left" Width="20%" CssClass="grid-header" VerticalAlign="Middle"></HeaderStyle>
														<ItemStyle Wrap="False" HorizontalAlign="Right" CssClass="grid-last-item"></ItemStyle>
														<ItemTemplate>
															<%# DataBinder.Eval(Container, "DataItem.CreatedDate") %>
														</ItemTemplate>
													</asp:TemplateColumn>
												</Columns>
												<PagerStyle Visible="False" HorizontalAlign="Center"></PagerStyle>
											</asp:datagrid></FONT><BR>
										<TABLE id="Table4" cellSpacing="0" cellPadding="0" width="100%" border="0">
											<TR>
												<TD style="FONT-SIZE: 11px" align="center">
													<asp:linkbutton id="lbFirstPage" onclick="PagerButtonClick" runat="server" CommandArgument="最前頁">最前頁</asp:linkbutton><FONT face="新細明體">&nbsp;&nbsp;
													</FONT>
													<asp:linkbutton id="lbLastPage" onclick="PagerButtonClick" runat="server" CommandArgument="最末頁">最末頁</asp:linkbutton><FONT face="新細明體">&nbsp;&nbsp; 
														&lt;&lt;&nbsp;&nbsp; 第 </FONT>
													<asp:label id="lblNowPage" runat="server"></asp:label><FONT face="新細明體">&nbsp;頁 , 共 </FONT>
													<asp:label id="lblTotalPage" runat="server"></asp:label><FONT face="新細明體">&nbsp;頁&nbsp;&nbsp;&nbsp;</FONT><FONT face="新細明體">
														&gt;&gt;&nbsp;&nbsp;</FONT>
													<asp:linkbutton id="lbPrevPage" onclick="PagerButtonClick" runat="server" CommandArgument="上一頁">上一頁</asp:linkbutton><FONT face="新細明體">&nbsp;&nbsp;
													</FONT>
													<asp:linkbutton id="lbNextPage" onclick="PagerButtonClick" runat="server" CommandArgument="下一頁">下一頁</asp:linkbutton><FONT face="新細明體"></FONT></TD>
												<TD style="FONT-SIZE: 11px" align="right"><FONT face="新細明體">&nbsp;總計
														<asp:Label id="lblRecordCount" runat="server"></asp:Label><FONT face="新細明體">&nbsp;筆&nbsp;
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
