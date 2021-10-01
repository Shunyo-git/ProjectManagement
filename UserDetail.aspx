<%@ Page Language="vb" AutoEventWireup="false" Codebehind="UserDetail.aspx.vb" Inherits="HsinYi.EIP.ProjectManagement.Web.UserDetail"%>
<%@ Register TagPrefix="uc1" TagName="myProjectTabs" Src="inc/mainTitleTabs.ascx" %>
<%@ Register TagPrefix="uc1" TagName="projectTabs" Src="inc/projectTabs.ascx" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>成員內容</title>
		<meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
		<meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
		<meta name="vs_defaultClientScript" content="JavaScript">
		<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
		<LINK href="css/Styles.css" type="text/css" rel="stylesheet">
		<script language="JavaScript" src="script/Global.js" type="text/JavaScript"></script>
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
<asp:LinkButton id="lbtnDelete" runat="server">刪除</asp:LinkButton>
															</asp:panel>
													</TD>
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
													<TD width="120" align="right" height="20" class="header-item">&nbsp;
														<asp:label id="userNamelbl" runat="server">姓名：</asp:label></TD>
													<TD class="item-text" height="20"><FONT face="新細明體">
															<asp:Label id="lblName" runat="server"></asp:Label></FONT></TD>
												</TR>
												<TR vAlign="top">
													<TD width="120" align="right" height="20" class="header-item">&nbsp;</TD>
													<TD class="item-text" height="20"><FONT face="新細明體"></FONT></TD>
												</TR>
												<TR>
													<TD width="120" align="right" height="20" class="header-item">&nbsp;
														<asp:label id="rolelbl" runat="server">角色：</asp:label></TD>
													<TD class="item-text" height="20"><FONT face="新細明體">
															<asp:Label id="lblRole" runat="server"></asp:Label></FONT></TD>
												</TR>
												<TR vAlign="top">
													<TD width="120" align="right" height="20" class="header-item">&nbsp;</TD>
													<TD class="item-text" height="20"><FONT face="新細明體"></FONT></TD>
												</TR>
												<TR>
													<TD width="120" align="right" height="20" class="header-item">&nbsp;
														<asp:label id="Label1" runat="server">組別：</asp:label></TD>
													<TD class="item-text" height="20">
														<asp:Label id="lblGroup" runat="server"></asp:Label></TD>
												</TR>
												<TR vAlign="top">
													<TD width="120" align="right" height="20" class="header-item">&nbsp;</TD>
													<TD class="item-text" height="20"><FONT face="新細明體"></FONT></TD>
												</TR>
												<TR>
													<TD align="right" width="120" height="20" class="header-item">
														<asp:Label id="Label5" runat="server">所屬單位：</asp:Label></TD>
													<TD class="item-text" height="20">
														<asp:label id="lblSourceUserDepartment" runat="server"></asp:label></TD>
												</TR>
												<TR vAlign="top">
													<TD align="right" width="120" height="20" class="header-item"></TD>
													<TD class="item-text" height="20"></TD>
												</TR>
												<TR>
													<TD align="right" width="120" height="20" class="header-item">
														<asp:label id="passwordlbl" runat="server">所屬職稱：</asp:label></TD>
													<TD class="item-text" height="20">
														<asp:Label id="lblSourceUserTitle" runat="server"></asp:Label></TD>
												</TR>
												<TR vAlign="top">
													<TD align="right" width="120" height="20" class="header-item"></TD>
													<TD class="item-text" height="20">&nbsp;</TD>
												</TR>
												<TR>
													<TD align="right" width="120" height="20" class="header-item">
														<asp:label id="lblTel" runat="server">聯絡資料：</asp:label></TD>
													<TD class="item-text" height="20">
														<asp:Label id="lblConnects" runat="server"></asp:Label></TD>
												</TR>
												<TR vAlign="top">
													<TD align="right" width="120" height="20" class="header-item"></TD>
													<TD class="item-text" height="20"><FONT face="新細明體"></FONT></TD>
												</TR>
												<TR>
													<TD vAlign="top" align="right" width="120" height="20" class="header-item">
														<asp:Label id="Label4" runat="server">工作簡述：</asp:Label></TD>
													<TD class="item-text" vAlign="top" height="20"><FONT face="新細明體">
															<asp:label id="lblDescription" runat="server"></asp:label></FONT></TD>
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
