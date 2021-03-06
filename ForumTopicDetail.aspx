<%@ Register TagPrefix="uc1" TagName="myProjectTabs" Src="inc/mainTitleTabs.ascx" %>
<%@ Register TagPrefix="uc1" TagName="projectTabs" Src="inc/projectTabs.ascx" %>
<%@ Page Language="vb" AutoEventWireup="false" Codebehind="ForumTopicDetail.aspx.vb" Inherits="HsinYi.EIP.ProjectManagement.Web.ForumTopicDetail"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>話題內容</title>
		<meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
		<meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
		<meta content="JavaScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
		<LINK href="css/Styles.css" type="text/css" rel="stylesheet">
		<script language="JavaScript" src="script/Global.js" type="text/JavaScript"></script>
	</HEAD>
	<body>
		<FONT face="新細明體">
			<FORM id="ProjectList" method="post" runat="server">
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
									<TD height="20"><FONT face="新細明體"><asp:panel id="PanelModify" runat="server">
												<TABLE id="Table6" cellSpacing="0" cellPadding="0" width="100%" border="0">
													<TR>
														<TD style="HEIGHT: 14px"><FONT face="新細明體"></FONT></TD>
														<TD style="HEIGHT: 14px" noWrap align="right"><IMG alt="" src="images/icon01.gif" align="absBottom">&nbsp;
															<asp:LinkButton id="lbtnNew" runat="server">新增主題</asp:LinkButton><IMG alt="" src="images/icon01.gif" align="absBottom">
															<asp:LinkButton id="lbtnReply" runat="server">新增回覆</asp:LinkButton>&nbsp;
														</TD>
													</TR>
													<TR>
														<TD noWrap align="right" colSpan="2"><IMG alt="" src="images/l_ba_bottom.GIF">
														</TD>
													</TR>
												</TABLE>
											</asp:panel></FONT></TD>
								</TR>
								<TR vAlign="top" height="*">
									<TD>
										<TABLE id="TableTopic" cellSpacing="0" cellPadding="0" width="100%" border="0">
											<TR>
												<TD class="forum-table-subject" style="HEIGHT: 18px"><asp:label id="lblSubject" runat="server">Subject</asp:label></TD>
											</TR>
											<TR>
												<TD class="admin-tan-border" style="PADDING-RIGHT: 10px; PADDING-LEFT: 18px; PADDING-BOTTOM: 5px; PADDING-TOP: 5px">
													<TABLE id="Table7" cellSpacing="0" cellPadding="0" width="100%" border="0">
														<TR>
															<TD class="forum-title-buttom" style="HEIGHT: 15px"><asp:label id="lblAuthor" runat="server">Author</asp:label></TD>
															<TD class="forum-title-buttom" style="HEIGHT: 12px" align="right"><asp:label id="lblModifiedDate" runat="server">CreatedDate</asp:label></TD>
														</TR>
														<TR>
															<TD class="forum-content" colSpan="2"><FONT face="新細明體"></FONT><FONT face="新細明體"><asp:label id="lblContent" runat="server">Content</asp:label></FONT></TD>
														</TR>
														<TR>
															<TD colSpan="2"><asp:label id="lblAttachment" runat="server"></asp:label></TD>
														</TR>
														<TR>
															<TD align="right" colSpan="2"><FONT face="新細明體"></FONT><IMG alt="" src="images/icon01.gif" align="absBottom">
																<asp:linkbutton id="lbtnEdit" runat="server">編輯</asp:linkbutton><IMG alt="" src="images/icon01.gif" align="absBottom">
																<asp:linkbutton id="lbtnDelete" runat="server">刪除</asp:linkbutton></TD>
														</TR>
													</TABLE>
												</TD>
											</TR>
										</TABLE>
										<p></p>
										<asp:datagrid id="dgReplys" runat="server" ShowHeader="False" AllowPaging="True" BorderColor="White"
											AutoGenerateColumns="False" CellPadding="2" Width="100%" BorderStyle="None">
											<Columns>
												<asp:TemplateColumn>
													<ItemTemplate>
														<TABLE id="Table9" cellSpacing="0" cellPadding="0" width="100%" border="0">
															<TR>
																<TD class="forum-reply" style="PADDING-RIGHT: 10px; PADDING-LEFT: 18px; PADDING-BOTTOM: 5px; PADDING-TOP: 5px">
																	<TABLE id="Table10" style="LINE-HEIGHT: 15pt" cellSpacing="0" cellPadding="0" width="100%"
																		border="0">
																		<TR>
																			<TD class="forum-title-buttom" style="HEIGHT: 15px">
																				<asp:Label id=Label6 runat="server" Text='<%# DataBinder.Eval(Container, "DataItem.CreatedUserName")  %>'>Author</asp:Label></TD>
																			<TD class="forum-title-buttom" style="HEIGHT: 12px" align="right">
																				<asp:Label id=Label5 runat="server" Text='<%# DataBinder.Eval(Container, "DataItem.CreatedDate")  %>'>CreatedDate</asp:Label></TD>
																		</TR>
																		<TR>
																			<TD class="forum-content" colSpan="2"><FONT face="新細明體"></FONT><FONT face="新細明體">
																					<asp:Label id=Label4 runat="server" Text='<%# DataBinder.Eval(Container, "DataItem.Content")  %>'>Content</asp:Label></FONT></TD>
																		</TR>
																		<TR>
																			<TD colSpan="2">
																				<%# showAttachmentTable( DataBinder.Eval(Container, "DataItem.DocumentID"))%>
																			</TD>
																		</TR>
																		<TR>
																			<TD align="right" colSpan="2"><FONT face="新細明體"></FONT><IMG alt="" src="images/icon01.gif" align="absBottom">
																				<A href='ForumTopicModify.aspx?projectID=<%# DataBinder.Eval(Container, "DataItem.ProjectID")  %>&amp;index=-1&amp;projectIndex=6&amp;TopicID=<%# DataBinder.Eval(Container, "DataItem.TopicID")  %>&amp;ReplyTo=<%# DataBinder.Eval(Container, "DataItem.ReplyTo")  %>'>
																					編輯</A><IMG alt="" src="images/icon01.gif" align="absBottom"> <A href='javascript:DeleteReply(<%# DataBinder.Eval(Container, "DataItem.ReplyTo")  %>,<%# DataBinder.Eval(Container, "DataItem.TopicID")  %>);'>
																					刪除</A></TD>
																		</TR>
																	</TABLE>
																</TD>
															</TR>
														</TABLE>
													</ItemTemplate>
												</asp:TemplateColumn>
											</Columns>
											<PagerStyle HorizontalAlign="Right" Position="TopAndBottom" PageButtonCount="5" CssClass="forum-table-subject"
												Mode="NumericPages"></PagerStyle>
										</asp:datagrid></TD>
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
			<iframe name="nosee" width="0" height="0"></iframe></FONT>
	</body>
</HTML>
