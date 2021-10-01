<%@ Page Language="vb" AutoEventWireup="false" Codebehind="ProjectDetails.aspx.vb" Inherits="HsinYi.EIP.ProjectManagement.Web.ProjectDetails"%>
<%@ Register TagPrefix="cc2" Namespace="HsinYiMemberPicker.HsinYi.EIP.Security.MemberPicker" Assembly="HsinYiMemberPicker" %>
<%@ Register TagPrefix="cc1" Namespace="DatePicker.iX.Controls" Assembly="DatePicker" %>
<%@ Register TagPrefix="uc1" TagName="myProjectTabs" Src="inc/mainTitleTabs.ascx" %>
<%@ Register TagPrefix="uc1" TagName="projectTabs" Src="inc/projectTabs.ascx" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>專案內容</title>
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
						<TD vAlign="bottom" bgColor="buttonface" height="46"><FONT face="新細明體"><uc1:myprojecttabs id="MyProjectTabs2" runat="server"></uc1:myprojecttabs></FONT></TD>
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
									<TD class="header-gray" vAlign="top" height="20"><asp:panel id="PanelModify" runat="server" Visible="False">
											<TABLE id="Table6" cellSpacing="0" cellPadding="0" width="100%" border="0">
												<TR>
													<TD>
														<asp:label id="ErrorMessage" runat="server" CssClass="standard-text" ForeColor="Red" enableviewstate="False"></asp:label></TD>
													<TD noWrap align="right"><IMG alt="" src="images/icon01.gif" align="absBottom">
														<asp:LinkButton id="lbtnEdit" runat="server">修改</asp:LinkButton>&nbsp;<IMG alt="" src="images/icon01.gif" align="absBottom">
														<asp:LinkButton id="lbtnDelete" runat="server">刪除</asp:LinkButton></TD>
												</TR>
												<TR>
													<TD noWrap align="right" colSpan="2"><IMG alt="" src="images/l_ba_bottom.GIF">
													</TD>
												</TR>
											</TABLE>
										</asp:panel></TD>
								</TR>
								<TR vAlign="top" height="*">
									<TD><asp:panel id="PanelView" runat="server">
											<TABLE id="TableView" cellSpacing="0" cellPadding="0" width="100%" border="0">
												<TR>
													<TD class="header-item" align="right" width="120" height="30">專案名稱：
													</TD>
													<TD class="header-text" height="30">
														<TABLE id="Table5" cellSpacing="0" cellPadding="0" width="100%" border="0">
															<TR>
																<TD class="item-text" width="100%" height="20">
																	<asp:label id="lblProjectName" runat="server" CssClass="item-text">Label</asp:label></TD>
																<TD class="header-item" noWrap align="right" height="20">專案狀態：
																	<asp:Label id="lblState" runat="server" CssClass="item-text">Label</asp:Label></TD>
															</TR>
														</TABLE>
													</TD>
												</TR>
												<TR>
													<TD class="header-item" vAlign="top" align="right" width="120" height="30"></TD>
													<TD class="header-text" vAlign="top" height="30"><FONT face="新細明體"></FONT></TD>
												</TR>
												<TR>
													<TD class="header-item" vAlign="top" align="right" width="120" height="30"><FONT face="新細明體">專案類型：</FONT></TD>
													<TD class="item-text" vAlign="top" height="30">
														<asp:Label id="lblType" runat="server">Label</asp:Label></TD>
												</TR>
												<TR>
													<TD class="header-item" vAlign="top" align="right" width="120" height="30"><FONT face="新細明體">專案期間：</FONT></TD>
													<TD class="item-text" vAlign="top" height="30">
														<asp:Label id="lblProjectTime" runat="server">Label</asp:Label></TD>
												</TR>
												<TR>
													<TD class="header-item" vAlign="top" align="right" width="120" height="30"><FONT face="新細明體">提案摘要：</FONT></TD>
													<TD class="item-text" vAlign="top" height="30">
														<asp:Label id="lblDescription" runat="server">Label</asp:Label></TD>
												</TR>
												<TR>
													<TD class="header-item" vAlign="top" align="right" width="120" height="30"></TD>
													<TD class="header-text" vAlign="top" height="30"></TD>
												</TR>
												<TR>
													<TD class="header-item" align="right" width="120" height="30"><FONT face="新細明體">專案負責人：</FONT></TD>
													<TD class="item-text" height="30">
														<asp:Label id="lblManager" runat="server">Label</asp:Label>
														<asp:Label id="Label1" runat="server" CssClass="header-item">　　　　　單位：</asp:Label>
														<asp:Label id="lblManagerDepartment" runat="server">Label</asp:Label></TD>
												</TR>
												<TR>
													<TD class="header-item" align="right" width="120" height="30"><FONT face="新細明體">主要執行者：</FONT></TD>
													<TD class="item-text" height="30">
														<asp:Label id="lblChief" runat="server">Label</asp:Label>
														<asp:Label id="Label3" runat="server" CssClass="header-item">　　　　　單位：</asp:Label>
														<asp:Label id="lblChiefDepartment" runat="server">Label</asp:Label></TD>
												</TR>
												<TR>
													<TD class="header-item" vAlign="top" align="right" width="120" height="30"></TD>
													<TD class="header-text" vAlign="top" height="30"><FONT face="新細明體"></FONT></TD>
												</TR>
												<TR>
													<TD class="header-item" vAlign="top" align="right" width="120" height="30"><FONT face="新細明體">專案補助政府機構：</FONT></TD>
													<TD class="item-text" vAlign="top" height="30">
														<asp:Label id="lblSupplementaryUnits" runat="server">Label</asp:Label></TD>
												</TR>
												<TR>
													<TD class="header-item" vAlign="top" align="right" width="120" height="30"><FONT face="新細明體">核定之補助經費：</FONT></TD>
													<TD class="item-text" vAlign="top" height="30">
														<asp:Label id="lblSupplementarySubsidy" runat="server">Label</asp:Label></TD>
												</TR>
												<TR>
													<TD class="header-item" vAlign="top" align="right" width="120" height="30"><FONT face="新細明體">協辦/贊助單位：</FONT></TD>
													<TD class="item-text" vAlign="top" height="30">
														<asp:Label id="lblAssistUnits" runat="server">Label</asp:Label></TD>
												</TR>
											</TABLE>
										</asp:panel></TD>
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
