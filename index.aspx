<%@ Register TagPrefix="uc1" TagName="myProjectTabs" Src="inc/mainTitleTabs.ascx" %>
<%@ Register TagPrefix="uc1" TagName="myForumTopic" Src="inc/myForumTopic.ascx" %>
<%@ Page  Culture="zh-tw" Language="vb" AutoEventWireup="false" Codebehind="index.aspx.vb" Inherits="HsinYi.EIP.ProjectManagement.Web.index"%>
<%@ Register TagPrefix="uc1" TagName="myCalendar" Src="inc/myCalendar.ascx" %>
<%@ Register TagPrefix="uc1" TagName="myProjectList" Src="inc/myProjectList.ascx" %>
<%@ Register TagPrefix="uc1" TagName="myDocuments" Src="inc/myDocuments.ascx" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>我的首頁</title>
		<meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
		<meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
		<meta content="JavaScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
		<LINK href="css/Styles.css" type="text/css" rel="stylesheet">
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
						<TD vAlign="top">
							<TABLE id="Table3" height="530" cellSpacing="20" cellPadding="0" width="100%" border="0">
								<TR vAlign="top" height="*">
									<TD>
										<TABLE class="tab-table" id="Table4" cellSpacing="5" cellPadding="0" width="100%" border="0">
											<TR>
												<TD class="admin-tan-border" vAlign="top" noWrap width="80%" height="150"><FONT class="item-text" face="新細明體">
														<uc1:myProjectList id="MyProjectList1" runat="server"></uc1:myProjectList></FONT></TD>
												<TD class="admin-tan-border" vAlign="top" height="150" rowSpan="3">
													<uc1:myCalendar id="MyCalendar1" runat="server"></uc1:myCalendar>
												</TD>
											</TR>
											<TR>
												<TD class="admin-tan-border" style="HEIGHT: 150px" vAlign="top" height="150"><FONT face="新細明體">&nbsp;
														<uc1:myDocuments id="MyDocuments1" runat="server"></uc1:myDocuments></FONT></TD>
											</TR>
											<TR>
												<TD class="admin-tan-border" vAlign="top" height="150"><FONT face="新細明體">
														<uc1:myForumTopic id="MyForumTopic1" runat="server"></uc1:myForumTopic></FONT></TD>
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
