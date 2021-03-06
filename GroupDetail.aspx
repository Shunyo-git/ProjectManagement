<%@ Register TagPrefix="uc1" TagName="projectTabs" Src="inc/projectTabs.ascx" %>
<%@ Register TagPrefix="uc1" TagName="myProjectTabs" Src="inc/mainTitleTabs.ascx" %>
<%@ Page Language="vb" AutoEventWireup="false" Codebehind="GroupDetail.aspx.vb" Inherits="HsinYi.EIP.ProjectManagement.Web.GroupDetail"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>工作群組內容</title>
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
						<TD vAlign="top" height="15"><IMG height="15" src="images/spacer.gif" width="15">
						</TD>
					</TR>
				</TABLE>
				<TABLE id="Table2" cellSpacing="0" cellPadding="0" width="100%" border="0">
					<TR>
						<TD width="8"><IMG height="8" src="images/spacer.gif" width="8"></TD>
						<TD vAlign="top"><uc1:projecttabs id="ProjectTabs1" runat="server"></uc1:projecttabs>
							<TABLE class="admin-tan-border" id="Table3" height="530" cellSpacing="20" cellPadding="0"
								width="100%" border="0">
								<TR>
									<TD height="20"><FONT face="新細明體">
											<TABLE id="Table6" cellSpacing="0" cellPadding="0" width="100%" border="0">
												<TR>
													<TD></TD>
													<TD noWrap align="right"><asp:panel id="PanelModify" runat="server" Visible="False"><IMG alt="" src="images/icon01.gif" align="absBottom"> 
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
													<TD class="header-item" align="right" width="120" height="20">&nbsp;
														<asp:label id="userNamelbl" runat="server">群組名稱：</asp:label></TD>
													<TD class="item-text" height="20"><FONT face="新細明體"><asp:label id="lblName" runat="server"></asp:label></FONT></TD>
												</TR>
												<TR vAlign="top">
													<TD class="header-item" align="right" width="120" height="20">&nbsp;</TD>
													<TD class="item-text" height="20"><FONT face="新細明體"></FONT></TD>
												</TR>
												<TR>
													<TD class="header-item" vAlign="top" align="right" width="120" height="20"><asp:label id="Label4" runat="server">敘述說明：</asp:label></TD>
													<TD class="item-text" vAlign="top" height="20"><FONT face="新細明體"><asp:label id="lblDescription" runat="server"></asp:label></FONT></TD>
												</TR>
												<TR>
													<TD class="header-item" align="right" width="120" height="20">&nbsp;</TD>
													<TD class="item-text" height="20">&nbsp;&nbsp;</TD>
												</TR>
												<TR>
													<TD class="header-item" vAlign="top" align="right" width="120" height="20"><asp:label id="Label1" runat="server">群組成員：</asp:label></TD>
													<TD class="item-text" vAlign="top" height="20"><FONT face="新細明體"><asp:label id="lblMembers" runat="server"></asp:label></FONT></TD>
												</TR>
												<TR>
													<TD class="header-item" align="right" width="120" height="20">&nbsp;</TD>
													<TD class="item-text" height="20">&nbsp;&nbsp;</TD>
												</TR>
												<TR vAlign="bottom">
													<TD class="header-item" align="right" width="120" height="20">&nbsp;</TD>
													<TD class="item-text" height="20"><IMG src="images/spacer.gif" width="8"></TD>
												</TR>
												<TR>
													<TD class="header-item" align="right" width="120" height="20"><IMG height="11" src="images/spacer.gif" width="11"></TD>
												</TR>
												<TR>
													<TD class="item-text" vAlign="top" align="right" width="120" colSpan="5" height="20"><IMG height="15" src="images/spacer.gif" width="15"></TD>
												</TR>
											</TABLE>
										</FONT>
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
