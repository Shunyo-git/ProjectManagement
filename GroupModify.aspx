<%@ Page Language="vb" AutoEventWireup="false" Codebehind="GroupModify.aspx.vb" Inherits="HsinYi.EIP.ProjectManagement.Web.GroupModify"%>
<%@ Register TagPrefix="uc1" TagName="projectTabs" Src="inc/projectTabs.ascx" %>
<%@ Register TagPrefix="uc1" TagName="myProjectTabs" Src="inc/mainTitleTabs.ascx" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>�s��u�@�s�R</title>
		<meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
		<meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
		<meta content="JavaScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
		<LINK href="css/Styles.css" type="text/css" rel="stylesheet">
		<script language="JavaScript" src="script/Global.js" type="text/JavaScript"></script>
	</HEAD>
	<body>
		<FONT face="�s�ө���">
			<FORM id="ProjectList" method="post" runat="server">
				<TABLE id="Table1" cellSpacing="0" cellPadding="0" width="100%" border="0">
					<TR>
						<TD vAlign="bottom" bgColor="buttonface" height="46"><FONT face="�s�ө���"><uc1:myprojecttabs id="MyProjectTabs1" runat="server"></uc1:myprojecttabs></FONT></TD>
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
									<TD height="20"><FONT face="�s�ө���">
											<asp:panel id="PanelModify" runat="server" Visible="False">
												<TABLE id="Table5" cellSpacing="0" cellPadding="0" width="100%" border="0">
													<TR>
														<TD>
															<asp:label id="ErrorMessage" runat="server" enableviewstate="False" ForeColor="Red" CssClass="standard-text"></asp:label></TD>
														<TD noWrap align="right">&nbsp;&nbsp; <IMG alt="" src="images/icon01.gif" align="absBottom">&nbsp;
															<asp:LinkButton id="lbtnSave" runat="server">�x�s</asp:LinkButton>&nbsp;<IMG alt="" src="images/icon01.gif" align="absBottom">
															<asp:LinkButton id="lbtnCancel" runat="server" CausesValidation="False">����</asp:LinkButton></TD>
													</TR>
													<TR>
														<TD noWrap align="right" colSpan="2"><IMG alt="" src="images/l_ba_bottom.GIF">
														</TD>
													</TR>
												</TABLE>
											</asp:panel>
										</FONT>
									</TD>
								</TR>
								<TR vAlign="top" height="*">
									<TD><FONT face="�s�ө���">
											<TABLE id="Table4" cellSpacing="0" cellPadding="0" width="100%" border="0">
												<TR>
													<TD class="header-item" align="right" width="120" height="20">&nbsp;
														<asp:label id="userNamelbl" runat="server">�s�զW�١G</asp:label></TD>
													<TD class="item-text" height="20"><FONT face="�s�ө���"><asp:textbox id="txtName" runat="server" CssClass="standard-text" MaxLength="50" Width="180px"></asp:textbox><asp:requiredfieldvalidator id="ProjectNameRequiredfieldvalidator" runat="server" ErrorMessage="�s�զW�٤��o�ť�." ControlToValidate="txtName"
																Display="Dynamic"></asp:requiredfieldvalidator></FONT></TD>
												</TR>
												<TR vAlign="top">
													<TD class="header-item" align="right" width="120" height="20">&nbsp;</TD>
													<TD class="item-text" height="20"><FONT face="�s�ө���"></FONT></TD>
												</TR>
												<TR>
													<TD class="header-item" vAlign="top" align="right" width="120" height="20"><asp:label id="Label4" runat="server">�ԭz�����G</asp:label></TD>
													<TD class="item-text" vAlign="top" height="20"><FONT face="�s�ө���"><asp:textbox id="txtDescription" runat="server" CssClass="standard-text" TextMode="MultiLine"
																Rows="5" MaxLength="50" Width="350px"></asp:textbox></FONT></TD>
												</TR>
												<TR>
													<TD class="header-item" align="right" width="120" height="20">&nbsp;</TD>
													<TD class="item-text" height="20">&nbsp;&nbsp;</TD>
												</TR>
												<TR>
													<TD class="header-item" vAlign="top" align="right" width="120" height="20"><asp:label id="Label1" runat="server">�s�զ����G</asp:label></TD>
													<TD class="item-text" vAlign="top" height="20"><asp:listbox id="Members" runat="server" Width="220px" DataValueField="UserID" DataTextField="UserName"
															SelectionMode="Multiple" Height="122px" rows="9" cssclass="box"></asp:listbox><FONT face="�s�ө���">(�h��ɽЫ���Ctrl��i���I��)</FONT>
													</TD>
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