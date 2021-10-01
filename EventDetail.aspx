<%@ Register TagPrefix="uc1" TagName="myProjectTabs" Src="inc/mainTitleTabs.ascx" %>
<%@ Register TagPrefix="uc1" TagName="projectTabs" Src="inc/projectTabs.ascx" %>
<%@ Page Language="vb" AutoEventWireup="false" Codebehind="EventDetail.aspx.vb" Inherits="HsinYi.EIP.ProjectManagement.Web.EventDetail"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>特殊事件內容</title>
		<meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
		<meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
		<meta name="vs_defaultClientScript" content="JavaScript">
		<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
		<LINK href="css/Styles.css" type="text/css" rel="stylesheet">
	</HEAD>
	<body>
		<FONT face="新細明體">
			<FORM id="ProjectList" method="post" runat="server">
				<TABLE id="Table1" cellSpacing="0" cellPadding="0" width="100%" border="0">
					<TR>
						<TD vAlign="bottom" bgColor="buttonface" height="46"><FONT face="新細明體">
								<uc1:myprojecttabs id="MyProjectTabs1" runat="server"></uc1:myprojecttabs></FONT></TD>
					</TR>
					<TR>
						<TD vAlign="top" height="15"><IMG height="15" src="images/spacer.gif" width="15">
						</TD>
					</TR>
				</TABLE>
				<TABLE id="Table2" cellSpacing="0" cellPadding="0" width="100%" border="0">
					<TR>
						<TD width="8"><IMG height="8" src="images/spacer.gif" width="8"></TD>
						<TD vAlign="top">
							<uc1:projecttabs id="ProjectTabs1" runat="server"></uc1:projecttabs>
							<TABLE class="admin-tan-border" id="Table3" height="530" cellSpacing="20" cellPadding="0"
								width="100%" border="0">
								<TR>
									<TD height="20"><FONT face="新細明體">
											<TABLE id="Table6" cellSpacing="0" cellPadding="0" width="100%" border="0">
												<TR>
													<TD></TD>
													<TD noWrap align="right">
														<asp:panel id="PanelModify" runat="server" Visible="False"><IMG alt="" src="images/icon01.gif" align="absBottom"> 
<asp:LinkButton id="lbtnEdit" runat="server">修改</asp:LinkButton>&nbsp;<IMG alt="" src="images/icon01.gif" align="absBottom"> 
<asp:LinkButton id="lbtnDelete" runat="server">刪除</asp:LinkButton></asp:panel></TD>
												</TR>
												<TR>
													<TD noWrap align="right" colSpan="2"><IMG alt="" src="images/l_ba_bottom.GIF">
													</TD>
												</TR>
											</TABLE>
										</FONT>
									</TD>
								</TR>
								<TR vAlign="top" height="*">
									<TD><FONT face="新細明體">
											<TABLE id="Table4" cellSpacing="0" cellPadding="0" width="100%" border="0">
												<TR>
													<TD class="header-item" align="right" width="120" height="20" vAlign="top">&nbsp;
														<asp:label id="itnTopic" runat="server">主題：</asp:label></TD>
													<TD class="item-text" height="20" vAlign="top"><FONT face="新細明體">
															<asp:Label id="lblTopic" runat="server"></asp:Label></FONT></TD>
												</TR>
												<TR vAlign="top">
													<TD class="header-item" align="right" width="120" height="20" vAlign="top">&nbsp;</TD>
													<TD class="item-text" height="20" vAlign="top"><FONT face="新細明體"></FONT></TD>
												</TR>
												<TR>
													<TD class="header-item" align="right" width="120" height="20" vAlign="top">&nbsp;
														<asp:label id="itnDate" runat="server">時間：</asp:label></TD>
													<TD class="item-text" height="20" vAlign="top"><FONT face="新細明體">
															<asp:Label id="lblEventDate" runat="server"></asp:Label></FONT></TD>
												</TR>
												<TR vAlign="top">
													<TD class="header-item" align="right" width="120" height="20" vAlign="top">&nbsp;</TD>
													<TD class="item-text" height="20" vAlign="top"><FONT face="新細明體"></FONT></TD>
												</TR>
												<TR>
													<TD class="header-item" align="right" width="120" height="20" vAlign="top">&nbsp;
														<asp:label id="Label1" runat="server">事件說明：</asp:label></TD>
													<TD class="item-text" height="20" vAlign="top">
														<asp:Label id="lblReason" runat="server"></asp:Label></TD>
												</TR>
												<TR vAlign="top">
													<TD class="header-item" align="right" width="120" height="20" vAlign="top"></TD>
													<TD class="item-text" height="20" vAlign="top"><FONT face="新細明體"></FONT></TD>
												</TR>
												<TR>
													<TD class="header-item" align="right" width="120" height="20" vAlign="top">
														<asp:label id="passwordlbl" runat="server">處理方法：</asp:label></TD>
													<TD class="item-text" height="20" vAlign="top">
														<asp:Label id="lblAnswer" runat="server"></asp:Label></TD>
												</TR>
												<TR vAlign="top">
													<TD class="header-item" align="right" width="120" height="20" vAlign="top"></TD>
													<TD class="item-text" height="20" vAlign="top">&nbsp;</TD>
												</TR>
												<TR>
													<TD class="header-item" align="right" width="120" height="20" vAlign="top">
														<asp:label id="lblTel" runat="server">處理成效：</asp:label></TD>
													<TD class="item-text" height="20" vAlign="top">
														<asp:Label id="lblEffect" runat="server"></asp:Label></TD>
												</TR>
												<TR vAlign="top">
													<TD class="header-item" align="right" width="120" height="20" vAlign="top"></TD>
													<TD class="item-text" height="20" vAlign="top"><FONT face="新細明體"></FONT></TD>
												</TR>
												<TR>
													<TD class="header-item" vAlign="top" align="right" width="120" height="20">
														<asp:Label id="Label4" runat="server">相關資料：</asp:Label></TD>
													<TD class="item-text" vAlign="top" height="20"><FONT face="新細明體">
															<asp:label id="lblAttachment" runat="server"></asp:label></FONT></TD>
												</TR>
												<TR>
													<TD class="header-item" align="right" width="120" height="20" vAlign="top">&nbsp;</TD>
													<TD class="item-text" height="20" vAlign="top">&nbsp;&nbsp;</TD>
												</TR>
												<TR vAlign="bottom">
													<TD class="header-item" align="right" width="120" height="20" vAlign="top">
														<asp:Label id="Label2" runat="server">新增者：</asp:Label>&nbsp;</TD>
													<TD class="item-text" height="20" vAlign="top">
														<asp:Label id="lblCreatedUserID" runat="server"></asp:Label><FONT face="新細明體">&nbsp; 
															( </FONT>
														<asp:Label id="lblCreatedDate" runat="server">Label</asp:Label><FONT face="新細明體">)</FONT></TD>
												</TR>
												<TR>
													<TD class="header-item" vAlign="top" align="right" width="120" height="20"></TD>
													<TD class="item-text" vAlign="top" height="20"></TD>
												</TR>
												<TR>
													<TD class="header-item" vAlign="top" align="right" width="120" height="20"><FONT face="新細明體">
															<asp:Label id="Label6" runat="server">最後修改：</asp:Label></FONT></TD>
													<TD class="item-text" vAlign="top" height="20">
														<asp:Label id="lblModifiedUserID" runat="server"></asp:Label><FONT face="新細明體">&nbsp; 
															( </FONT>
														<asp:Label id="lblModifiedDate" runat="server">Label</asp:Label><FONT face="新細明體">)</FONT></TD>
												</TR>
												<TR>
													<TD class="header-item" vAlign="top" align="right" width="120" height="20"></TD>
													<TD class="item-text" vAlign="top" height="20"><IMG src="images/spacer.gif" width="8"></TD>
												</TR>
											</TABLE>
										</FONT>
									</TD>
								</TR>
							</TABLE>
						</TD>
						<TD width="11"><IMG height="11" src="images/spacer.gif" width="11"></TD>
					</TR>
				</TABLE>
			</FORM>
			<iframe name="nosee" width="0" height="0"></iframe></FONT>
	</body>
</HTML>
