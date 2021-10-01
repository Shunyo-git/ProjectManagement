<%@ Register TagPrefix="cc2" Namespace="HsinYiMemberPicker.HsinYi.EIP.Security.MemberPicker" Assembly="HsinYiMemberPicker" %>
<%@ Register TagPrefix="uc1" TagName="projectTabs" Src="inc/projectTabs.ascx" %>
<%@ Register TagPrefix="uc1" TagName="myProjectTabs" Src="inc/mainTitleTabs.ascx" %>
<%@ Page Language="vb" AutoEventWireup="false" Codebehind="UserModify.aspx.vb" Inherits="HsinYi.EIP.ProjectManagement.Web.UserModify"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>編輯成員</title>
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
						<TD vAlign="bottom" bgColor="buttonface" height="46"><FONT face="新細明體"><uc1:myprojecttabs id="MyProjectTabs1" runat="server"></uc1:myprojecttabs></FONT></TD>
					</TR>
					<TR>
						<TD vAlign="top" height="15"><IMG height="15" src="images/spacer.gif" width="15">
						</TD>
					</TR>
				</TABLE>
				<TABLE id="Table2" cellSpacing="0" cellPadding="0" width="100%" border="0">
					<TBODY>
						<TR>
							<TD width="8"><IMG height="8" src="images/spacer.gif" width="8"></TD>
							<TD vAlign="top"><uc1:projecttabs id="ProjectTabs1" runat="server"></uc1:projecttabs>
								<TABLE class="admin-tan-border" id="Table3" height="530" cellSpacing="20" cellPadding="0"
									width="100%" border="0">
									<TR>
										<TD height="20"><asp:panel id="PanelModify" runat="server" Visible="False">
												<TABLE id="Table6" cellSpacing="0" cellPadding="0" width="100%" border="0">
													<TR>
														<TD>
															<asp:label id="ErrorMessage" runat="server" CssClass="standard-text" ForeColor="Red" enableviewstate="False"></asp:label></TD>
														<TD noWrap align="right">&nbsp;&nbsp; <IMG alt="" src="images/icon01.gif" align="absBottom">&nbsp;
															<asp:LinkButton id="lbtnSave" runat="server">儲存</asp:LinkButton>&nbsp;<IMG alt="" src="images/icon01.gif" align="absBottom">
															<asp:LinkButton id="lbtnCancel" runat="server" CausesValidation="False">取消</asp:LinkButton></TD>
													</TR>
													<TR>
														<TD noWrap align="right" colSpan="2"><IMG alt="" src="images/l_ba_bottom.GIF">
														</TD>
													</TR>
												</TABLE>
											</asp:panel></TD>
									</TR>
									<TR vAlign="top" height="*">
										<TD><FONT face="新細明體">
												<TABLE id="Table4" height="500" cellSpacing="0" cellPadding="0" width="100%" border="0">
													<TBODY>
														<TR>
															<TD style="WIDTH: 119px" align="right" width="119" height="25">&nbsp;<asp:label id="userNamelbl" runat="server">姓名：</asp:label></TD>
															<TD height="25"><cc2:memberpicker id="txtName" runat="server" SelectUserOnly="True" ValidatorErrorMessage="姓名不得空白"
																	EnableEmpty="False"></cc2:memberpicker></TD>
														</TR>
														<TR>
															<TD align="right" height="25"><asp:label id="rolelbl" runat="server">角色：</asp:label></TD>
															<TD height="25"><asp:dropdownlist id="ddlRole" runat="server" CssClass="standard-text" Width="180px"></asp:dropdownlist></TD>
														</TR>
														<TR>
															<TD style="WIDTH: 119px" align="right" width="119" height="25">&nbsp;<asp:label id="Label1" runat="server">組別：</asp:label></TD>
															<TD height="25"><asp:dropdownlist id="ddlGroup" runat="server" CssClass="standard-text" Width="180px"></asp:dropdownlist></TD>
														</TR>
														<TR>
															<TD style="WIDTH: 119px" align="right" width="119" height="25">&nbsp;<asp:label id="lblSourceUserDepartment" runat="server">所屬單位：</asp:label></TD>
															<TD height="25"><asp:textbox id="txtSourceUserDepartment2" runat="server" CssClass="standard-text" MaxLength="50"
																	Width="180px" Visible="False"></asp:textbox>
																<cc2:memberpicker id="txtSourceUserDepartment" runat="server" EnableEmpty="False" ValidatorErrorMessage="姓名不得空白"
																	SelectUserOnly="False" SelectGroupOnly="True"></cc2:memberpicker></TD>
														</TR>
														<TR>
															<TD style="WIDTH: 119px" align="right" width="119" height="25">&nbsp;<asp:label id="passwordlbl" runat="server">所屬職稱：</asp:label></TD>
															<TD height="25"><asp:textbox id="txtSourceUserTitle" runat="server" CssClass="standard-text" MaxLength="50" Width="180px"></asp:textbox></TD>
														</TR>
														<TR>
															<TD vAlign="top" align="right" width="120" height="25">&nbsp;<asp:label id="lblTel" runat="server">聯絡資料：<BR> (非上班時間的聯絡方式)</asp:label></TD>
															<TD vAlign="top" height="25"><asp:textbox id="txtConnects" runat="server" CssClass="standard-text" MaxLength="100" Width="180px"></asp:textbox></TD>
														</TR>
														<TR>
															<TD style="WIDTH: 119px" vAlign="top" align="right" width="119" height="25"><asp:label id="lblDescription" runat="server">工作簡述：</asp:label></TD>
															<TD vAlign="top" height="25"><asp:textbox id="txtDescription" runat="server" CssClass="standard-text" MaxLength="50" Width="300px"
																	TextMode="MultiLine" Rows="3"></asp:textbox></TD>
														</TR>
														<TR>
															<TD style="WIDTH: 119px" align="right" width="119" height="25"><IMG height="11" src="images/spacer.gif" width="11"></TD>
														</TR>
														<TR>
															<TD vAlign="top" align="right" width="120" colSpan="5" height="25"><IMG height="15" src="images/spacer.gif" width="15"></TD>
														</TR>
													</TBODY>
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
					</TBODY>
				</TABLE>
			</FORM>
		</FONT>
	</body>
</HTML>
