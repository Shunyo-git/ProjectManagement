<%@ Page Language="vb" AutoEventWireup="false" Codebehind="ProjectGroup.aspx.vb" Inherits="HsinYi.EIP.ProjectManagement.Web.ProjectGroup"%>
<%@ Register TagPrefix="uc1" TagName="projectTabs" Src="inc/projectTabs.ascx" %>
<%@ Register TagPrefix="uc1" TagName="myProjectTabs" Src="inc/mainTitleTabs.ascx" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>工作群組</title>
		<meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
		<meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
		<meta content="JavaScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
		<LINK href="css/Styles.css" type="text/css" rel="stylesheet">
	</HEAD>
	<body>
		<FONT face="新細明體">
			<FORM id="GroupList" method="post" runat="server">
				<TABLE id="Table1" cellSpacing="0" cellPadding="0" width="100%" border="0">
					<TR>
						<TD vAlign="bottom" bgColor="buttonface" height="46"><FONT face="新細明體"><uc1:myprojecttabs id="MyProjectTabs1" runat="server"></uc1:myprojecttabs></FONT></TD>
					</TR>
					<TR>
						<TD vAlign="top" height="15"><IMG height="15" src="images/spacer.gif" width="15"></TD>
					</TR>
				</TABLE>
				<TABLE id="Table2" cellSpacing="0" cellPadding="0" width="100%" border="0">
					<TBODY>
						<TR>
							<TD width="8"><IMG height="8" src="images/spacer.gif" width="8"></TD>
							<TD vAlign="top"><uc1:projecttabs id="ProjectTabs1" runat="server"></uc1:projecttabs>
								<TABLE class="admin-tan-border" id="Table3" height="530" cellSpacing="20" cellPadding="0"
									width="100%" border="0">
									<TBODY>
										<TR>
											<TD height="20"><asp:panel id="PanelModify" runat="server">
													<TABLE id="Table6" cellSpacing="0" cellPadding="0" width="100%" border="0">
														<TR>
															<TD style="HEIGHT: 14px"></TD>
															<TD style="HEIGHT: 14px" noWrap align="right"><IMG alt="" src="images/icon01.gif" align="absBottom">&nbsp;
																<asp:LinkButton id="lbtnNew" runat="server">新增群組</asp:LinkButton>&nbsp;
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
											<TD><FONT face="新細明體"><asp:datagrid id="dgGroups" runat="server" AllowPaging="True" BorderColor="White" AllowSorting="True"
														AutoGenerateColumns="False" CellPadding="2" Width="100%" BorderStyle="None">
														<HeaderStyle Font-Bold="True" CssClass="grid-header"></HeaderStyle>
														<Columns>
															<asp:TemplateColumn SortExpression="Name" HeaderText="群組名稱">
																<HeaderStyle HorizontalAlign="Left" Width="20%" CssClass="grid-header" VerticalAlign="Middle"></HeaderStyle>
																<ItemStyle CssClass="grid-first-item"></ItemStyle>
																<ItemTemplate>
																	&nbsp;<a title="Group Detail" class="blacklineheight" href='GroupDetail.aspx?GroupID=<%#DataBinder.Eval(Container.DataItem, "GroupID")%>&ProjectID=<%#DataBinder.Eval(Container.DataItem, "ProjectID")%>&index=-1&projectIndex=1'>
																		<asp:label ID="Label1" Text='<%# DataBinder.Eval(Container, "DataItem.Name") %>' Runat="server" Visible="true" />
																	</a>
																</ItemTemplate>
															</asp:TemplateColumn>
															<asp:TemplateColumn SortExpression="Description" HeaderText="描述說明">
																<HeaderStyle HorizontalAlign="Left" Width="40%" CssClass="grid-header" VerticalAlign="Middle"></HeaderStyle>
																<ItemStyle CssClass="grid-item"></ItemStyle>
																<ItemTemplate>
																	&nbsp;
																	<asp:label ID="Label3" Text='<%# DataBinder.Eval(Container, "DataItem.Description") %>' Runat="server" Visible="true" />
																</ItemTemplate>
															</asp:TemplateColumn>
															<asp:TemplateColumn SortExpression="Description" HeaderText="組員">
																<HeaderStyle HorizontalAlign="Left" Width="40%" CssClass="grid-header" VerticalAlign="Middle"></HeaderStyle>
																<ItemStyle CssClass="grid-last-item"></ItemStyle>
																<ItemTemplate>
																	&nbsp;
																	<asp:label ID="Label2" Text='<%# getGroupMembers(DataBinder.Eval(Container, "DataItem.GroupID")) %>' Runat="server" Visible="true" />
																</ItemTemplate>
															</asp:TemplateColumn>
														</Columns>
														<PagerStyle Visible="False" HorizontalAlign="Center"></PagerStyle>
													</asp:datagrid></FONT>
												<br>
												<TABLE id="Table4" cellSpacing="0" cellPadding="0" width="100%" border="0">
													<TBODY>
														<TR>
															<TD align="center" style="FONT-SIZE: 11px">
																<asp:linkbutton id="lbFirstPage" runat="server" OnClick="PagerButtonClick" CommandArgument="最前頁">最前頁</asp:linkbutton><FONT face="新細明體">&nbsp;&nbsp;
																</FONT>
																<asp:linkbutton id="lbLastPage" onclick="PagerButtonClick" runat="server" CommandArgument="最末頁">最末頁</asp:linkbutton><FONT face="新細明體">&nbsp;&nbsp; 
																	&lt;&lt;&nbsp;&nbsp; 第 </FONT>
																<asp:label id="lblNowPage" runat="server"></asp:label><FONT face="新細明體">&nbsp;頁 , 共 </FONT>
																<asp:label id="lblTotalPage" runat="server"></asp:label><FONT face="新細明體"> &nbsp;頁&nbsp;&nbsp;&nbsp;</FONT><FONT face="新細明體">
																	&gt;&gt;&nbsp;&nbsp;</FONT>
																<asp:linkbutton id="lbPrevPage" onclick="PagerButtonClick" runat="server" CommandArgument="上一頁">上一頁</asp:linkbutton><FONT face="新細明體">&nbsp;&nbsp;
																</FONT>
																<asp:linkbutton id="lbNextPage" runat="server" OnClick="PagerButtonClick" CommandArgument="下一頁">下一頁</asp:linkbutton><FONT face="新細明體"></FONT>
															</TD>
															<TD style="FONT-SIZE: 11px" align="right"><FONT face="新細明體">&nbsp;總計
																	<asp:Label id="lblRecordCount" runat="server"></asp:Label><FONT face="新細明體">&nbsp;筆&nbsp;
																	</FONT></FONT>
															</TD>
														</TR>
													</TBODY>
												</TABLE>
											</TD>
										</TR>
										<TR vAlign="top">
											<td height="12" colSpan="3"></td>
										</TR>
									</TBODY>
								</TABLE>
							</TD>
							<TD width="11"><IMG height="11" src="images/spacer.gif" width="11"></TD>
						</TR>
						<TR>
							<TD vAlign="top" colSpan="5" height="15"><IMG height="15" src="images/spacer.gif" width="15"></TD>
						</TR>
					</TBODY>
				</TABLE>
			</FORM>
		</FONT></TD></TR></TBODY></TABLE></TD></TR></TBODY></TABLE></TD></TR></TBODY></TABLE></FORM></FONT>
	</body>
</HTML>
