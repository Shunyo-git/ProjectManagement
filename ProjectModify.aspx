<%@ Register TagPrefix="cc2" Namespace="HsinYiMemberPicker.HsinYi.EIP.Security.MemberPicker" Assembly="HsinYiMemberPicker" %>
<%@ Register TagPrefix="cc1" Namespace="DatePicker.iX.Controls" Assembly="DatePicker" %>
<%@ Register TagPrefix="uc1" TagName="projectTabs" Src="inc/projectTabs.ascx" %>
<%@ Register TagPrefix="uc1" TagName="myProjectTabs" Src="inc/mainTitleTabs.ascx" %>
<%@ Page Language="vb" AutoEventWireup="false" Codebehind="ProjectModify.aspx.vb" Inherits="HsinYi.EIP.ProjectManagement.BusinessLogicLayer.ProjectModify"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>編輯專案</title>
		<meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
		<meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
		<meta name="vs_defaultClientScript" content="JavaScript">
		<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
		<LINK href="css/Styles.css" type="text/css" rel="stylesheet">
	</HEAD>
	<body>
		<FONT face="新細明體">
			<FORM id="ProjectList" method="post" runat="server">
				<TABLE id="Table1" cellSpacing="0" cellPadding="0" width="100%" border="0">
					<TR>
						<TD vAlign="bottom" bgColor="buttonface" height="46"><FONT face="新細明體">
								<uc1:myProjectTabs id="MyProjectTabs1" runat="server"></uc1:myProjectTabs></FONT></TD>
					</TR>
					<TR>
						<TD vAlign="top" height="15"><IMG height="15" src="images/spacer.gif" width="15"></TD>
					</TR>
				</TABLE>
				<TABLE id="Table2" cellSpacing="0" cellPadding="0" width="100%" border="0">
					<TR>
						<TD width="8"><IMG height="8" src="images/spacer.gif" width="8"></TD>
						<TD vAlign="top">
							<uc1:projectTabs id="ProjectTabs1" runat="server"></uc1:projectTabs>
							<TABLE class="admin-tan-border" id="Table3" height="530" cellSpacing="20" cellPadding="0"
								width="100%" border="0">
								<TR>
									<TD class="header-gray" vAlign="top" height="20">
										<asp:panel id="PanelModify" runat="server">
											<TABLE id="Table6" cellSpacing="0" cellPadding="0" width="100%" border="0">
												<TR>
													<TD>
														<asp:label id="ErrorMessage" runat="server" CssClass="standard-text" ForeColor="Red" enableviewstate="False"></asp:label></TD>
													<TD noWrap align="right"><IMG alt="" src="images/icon01.gif" align="absBottom">&nbsp;
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
									<TD>
										<asp:panel id="PanelEdit" runat="server">
											<TABLE id="TableEdit" cellSpacing="0" cellPadding="0" width="100%" border="0">
												<TR>
													<TD class="header-item" align="right" width="120" height="25"><FONT face="新細明體">專案名稱：</FONT></TD>
													<TD class="header-text" height="25">
														<TABLE id="Table4" cellSpacing="0" cellPadding="0" width="100%" border="0">
															<TR>
																<TD class="header-item" width="100%" height="20">
																	<asp:TextBox id="txtProjectName" runat="server" Width="300px"></asp:TextBox>
																	<asp:requiredfieldvalidator id="ProjectNameRequiredfieldvalidator" runat="server" ErrorMessage="專案名稱不得空白." display="Dynamic"
																		ControlToValidate="txtProjectName"></asp:requiredfieldvalidator></TD>
																<TD class="header-item" noWrap align="right" height="20">專案狀態：
																	<asp:DropDownList id="ddlState" runat="server"></asp:DropDownList></TD>
															</TR>
														</TABLE>
													</TD>
												</TR>
												<TR>
													<TD class="header-item" align="right" width="120" height="25"></TD>
													<TD class="header-text" height="25"><FONT face="新細明體"></FONT></TD>
												</TR>
												<TR>
													<TD class="header-item" align="right" width="120" height="25"><FONT face="新細明體">專案類型：</FONT></TD>
													<TD class="header-text" height="25"><FONT face="新細明體">
															<asp:DropDownList id="ddlType" runat="server"></asp:DropDownList></FONT></TD>
												</TR>
												<TR>
													<TD class="header-item" align="right" width="120" height="25"><FONT face="新細明體">專案期間：</FONT></TD>
													<TD class="header-text" height="25">
														<cc1:DatePicker id="txtStartDate" runat="server" ControlCssClass="DatePicker" DateType="yyyy/mm/dd"
															imgDirectory="/app/outerapp/ProjectManagement/datePicker/"></cc1:DatePicker>～
														<cc1:DatePicker id="txtEndDate" runat="server" ControlCssClass="DatePicker" DateType="yyyy/mm/dd"
															imgDirectory="/app/outerapp/ProjectManagement/datePicker/"></cc1:DatePicker></TD>
												</TR>
												<TR>
													<TD class="header-item" vAlign="top" align="right" width="120" height="25"><FONT face="新細明體">提案摘要：</FONT></TD>
													<TD class="header-text" height="25">
														<asp:TextBox id="txtDescription" runat="server" Width="450px" TextMode="MultiLine" Rows="5"></asp:TextBox></TD>
												</TR>
												<TR>
													<TD class="header-item" align="right" width="120" height="25"></TD>
													<TD class="header-text" height="25"></TD>
												</TR>
												<TR>
													<TD class="header-item" style="HEIGHT: 25px" align="right" width="120" height="25"><FONT face="新細明體">專案負責人：</FONT></TD>
													<TD class="header-text" style="HEIGHT: 25px" height="25">
														<cc2:MemberPicker id="txtManager" runat="server" SelectUserOnly="True"></cc2:MemberPicker></TD>
												</TR>
												<TR>
													<TD class="header-item" align="right" width="120" height="25"><FONT face="新細明體">主要執行者：</FONT></TD>
													<TD class="header-text" height="25"><FONT face="新細明體">
															<cc2:MemberPicker id="txtChief" runat="server" SelectUserOnly="True"></cc2:MemberPicker></FONT></TD>
												</TR>
												<TR>
													<TD class="header-item" align="right" width="120" height="25"></TD>
													<TD class="header-text" height="25"><FONT face="新細明體"></FONT></TD>
												</TR>
												<TR>
													<TD class="header-item" vAlign="top" align="right" width="120" height="25"><FONT face="新細明體">專案補助政府機構：</FONT></TD>
													<TD class="header-text" height="25">
														<asp:TextBox id="txtSupplementaryUnits" runat="server" Width="450px" TextMode="MultiLine" Rows="3"></asp:TextBox></TD>
												</TR>
												<TR>
													<TD class="header-item" vAlign="top" align="right" width="120" height="25"><FONT face="新細明體">核定之補助經費：</FONT></TD>
													<TD class="header-text" height="25"><FONT face="新細明體">
															<asp:TextBox id="txtSupplementarySubsidy" runat="server" Width="450px" TextMode="MultiLine" Rows="3"></asp:TextBox></FONT></TD>
												</TR>
												<TR>
													<TD class="header-item" vAlign="top" align="right" width="120" height="25"><FONT face="新細明體">協辦/贊助單位：</FONT></TD>
													<TD class="header-text" height="25">
														<asp:TextBox id="txtAssistUnits" runat="server" Width="450px" TextMode="MultiLine" Rows="3"></asp:TextBox></TD>
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
