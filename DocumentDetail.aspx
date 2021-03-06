<%@ Page Language="vb" AutoEventWireup="false" Codebehind="DocumentDetail.aspx.vb" Inherits="HsinYi.EIP.ProjectManagement.Web.DocumentDetail"%>
<%@ Register TagPrefix="uc1" TagName="projectTabs" Src="inc/projectTabs.ascx" %>
<%@ Register TagPrefix="uc1" TagName="myProjectTabs" Src="inc/mainTitleTabs.ascx" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>文件內容</title>
		<meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
		<meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
		<meta content="JavaScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
		<LINK href="css/Styles.css" type="text/css" rel="stylesheet">
		<script language="JavaScript" src="script/Global.js" type="text/JavaScript"></script>
		<script language="javascript">
		function toClear(e) {
			//刪除檔案後,清除檔案顯示相關資料
			var obj=document.getElementById("txtFilename") ;
			obj.value=obj.value.replace(e.file,',');
			obj.value=obj.value.replace(',,',',');
			obj.value=obj.value.replace(',,',',');
			if (obj.value.substr(0,1)==',')
				obj.value=obj.value.substr(1);
			if (obj.value.substr(obj.value.length-1,1)==',')
				obj.value=obj.value.substr(0,obj.value.length-1);
			var fobj=document.all['f_' + e.file]
			fobj.style.display='none';
		}
		</script>
	</HEAD>
	<body>
		<FONT face="新細明體">
			<FORM id="DocumentDetail" method="post" runat="server">
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
											<TD>
												<TABLE id="Table4" cellSpacing="0" cellPadding="0" width="100%" border="0">
													<TBODY>
														<TR>
															<TD class="header-item" align="right" width="120" height="20">&nbsp;
																<asp:label id="itnTopic" runat="server">標題：</asp:label></TD>
															<TD class="item-text" height="20"><FONT face="新細明體">
																	<asp:Label id="lblTitle" runat="server"></asp:Label></FONT></TD>
														</TR>
														<TR vAlign="top">
															<TD class="header-item" vAlign="top" align="right" width="120" height="20">&nbsp;</TD>
															<TD class="item-text" vAlign="top" height="20"><FONT face="新細明體"></FONT></TD>
														</TR>
														<TR>
															<TD class="header-item" align="right" width="120" height="20"><asp:label id="Label2" runat="server">類別：</asp:label>&nbsp;
															</TD>
															<TD class="item-text" height="20"><FONT face="新細明體">
																	<asp:Label id="lblCategory" runat="server"></asp:Label></FONT></TD>
														</TR>
														<TR vAlign="top">
															<TD class="header-item" vAlign="top" align="right" width="120" height="20">&nbsp;</TD>
															<TD class="item-text" vAlign="top" height="20"><FONT face="新細明體"></FONT></TD>
														</TR>
														<TR>
															<TD class="header-item" vAlign="top" align="right" width="120" height="20">
																<asp:label id="Label1" runat="server">備註：</asp:label></TD>
															<TD class="item-text" vAlign="top" height="20">
																<asp:Label id="lblDescription" runat="server"></asp:Label></TD>
														</TR>
														<TR vAlign="top">
															<TD class="header-item" vAlign="top" align="right" width="120" height="20"></TD>
															<TD class="item-text" vAlign="top" height="20"><FONT face="新細明體"></FONT></TD>
														</TR>
														<TR>
															<TD class="header-item" vAlign="top" align="right" width="120" height="20"><asp:label id="Label4" runat="server">相關資料：</asp:label></TD>
															<TD class="item-text" vAlign="top" height="20"><asp:label id="lblAttachment" runat="server"></asp:label>
															</TD>
														</TR>
														<TR>
															<TD class="header-item" vAlign="top" align="right" width="120" height="20">&nbsp;</TD>
															<TD class="item-text" vAlign="top" height="20">&nbsp;&nbsp;</TD>
														</TR>
														<TR vAlign="bottom">
															<TD class="header-item" vAlign="top" align="right" width="120" height="20">
																<asp:Label id="Label3" runat="server">新增者：</asp:Label>&nbsp;</TD>
															<TD class="item-text" vAlign="top" height="20"><FONT face="新細明體">
																	<asp:Label id="lblCreatedUserID" runat="server"></asp:Label><FONT face="新細明體">&nbsp;
																	</FONT>(
																	<asp:Label id="lblCreatedDate" runat="server">Label</asp:Label><FONT face="新細明體">)</FONT></FONT></TD>
														</TR>
														<TR>
															<TD class="header-item" vAlign="top" align="right" width="120" height="20"></TD>
															<TD class="item-text" vAlign="top" height="20"><FONT face="新細明體"></FONT></TD>
														</TR>
														<TR>
															<TD class="header-item" vAlign="top" align="right" width="120" height="20"><FONT face="新細明體">
																	<asp:Label id="Label6" runat="server">最後修改：</asp:Label></FONT></TD>
															<TD class="item-text" vAlign="top" height="20"><FONT face="新細明體">
																	<asp:Label id="lblModifiedUserID" runat="server"></asp:Label>&nbsp; (
																	<asp:Label id="lblModifiedDate" runat="server">Label</asp:Label><FONT face="新細明體">)</FONT></FONT></TD>
															<TD class="item-text" vAlign="top" height="20"></TD>
														</TR>
														<TR>
															<TD class="header-item" vAlign="top" height="20" width="120" align="right"></TD>
															<TD class="item-text" vAlign="top" height="20">
															</TD>
														</TR>
														<TR>
															<TD class="header-item" vAlign="top" align="right" width="120" height="20"></TD>
															<TD class="item-text" vAlign="top" height="20"><IMG src="images/spacer.gif" width="8"></TD>
														</TR>
													</TBODY>
												</TABLE>
											</TD>
										</TR>
										<TR vAlign="top">
											<TD colSpan="3" height="12"></TD>
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
			<iframe name="nosee" width="0" height="0"></iframe></FONT>
	</body>
</HTML>
