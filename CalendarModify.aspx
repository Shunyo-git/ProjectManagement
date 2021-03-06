<%@ Register TagPrefix="uc1" TagName="myProjectTabs" Src="inc/mainTitleTabs.ascx" %>
<%@ Register TagPrefix="uc1" TagName="projectTabs" Src="inc/projectTabs.ascx" %>
<%@ Register TagPrefix="cc1" Namespace="DatePicker.iX.Controls" Assembly="DatePicker" %>
<%@ Page Language="vb"    AutoEventWireup="false" Codebehind="CalendarModify.aspx.vb" Inherits="HsinYi.EIP.ProjectManagement.Web.CalendarModify"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>編輯行事曆</title>
		<meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
		<meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
		<meta content="JavaScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
		<LINK href="css/Styles.css" type="text/css" rel="stylesheet">
	</HEAD>
	<body>
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
				<TBODY>
					<TR>
						<TD width="8"><IMG height="8" src="images/spacer.gif" width="8"></TD>
						<TD vAlign="top"><uc1:projecttabs id="ProjectTabs1" runat="server"></uc1:projecttabs>
							<TABLE class="admin-tan-border" id="Table3" height="530" cellSpacing="20" cellPadding="0"
								width="100%" border="0">
								<TBODY>
									<TR>
										<TD height="20"><asp:panel id="PanelModify" runat="server">
												<TABLE id="Table5" cellSpacing="0" cellPadding="0" width="100%" border="0">
													<TR>
														<TD>
															<asp:label id="ErrorMessage" runat="server" CssClass="standard-text" enableviewstate="False"
																ForeColor="Red"></asp:label></TD>
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
										<TD>
											<TABLE id="Table4" cellSpacing="0" cellPadding="0" width="100%" border="0">
												<TBODY>
													<TR>
														<TD class="header-item" style="HEIGHT: 21px" align="right" width="120" height="21">&nbsp;
															<asp:label id="itnTopic" runat="server">事項：</asp:label></TD>
														<TD class="item-text" style="HEIGHT: 21px" height="21"><FONT face="新細明體"><asp:textbox id="txtSubject" runat="server" Width="350px"></asp:textbox><asp:requiredfieldvalidator id="SubjectRequiredfieldvalidator" runat="server" ErrorMessage="事項不得空白." ControlToValidate="txtSubject"
																	Display="Dynamic"></asp:requiredfieldvalidator></FONT></TD>
													</TR>
													<TR>
														<TD class="header-item" style="HEIGHT: 20px" vAlign="top" align="right" width="120"
															height="20"></TD>
														<TD class="item-text" style="HEIGHT: 20px" vAlign="top" height="20"><FONT face="新細明體"></FONT></TD>
													</TR>
													<TR>
														<TD class="header-item" style="HEIGHT: 3px" vAlign="top" align="right" width="120" height="3"><FONT face="新細明體"><asp:label id="Label2" runat="server">類型：</asp:label></FONT></TD>
														<TD class="item-text" style="HEIGHT: 3px" vAlign="top" height="3"><asp:dropdownlist id="ddlAppointType" runat="server"></asp:dropdownlist></TD>
													</TR>
													<TR>
														<TD class="header-item" style="HEIGHT: 20px" vAlign="top" align="right" width="120"
															height="20"></TD>
														<TD class="item-text" style="HEIGHT: 20px" vAlign="top" height="20"></TD>
													</TR>
													<TR>
														<TD class="header-item" style="HEIGHT: 21px" vAlign="top" align="right" width="120"
															height="21"><FONT face="新細明體"><asp:label id="Label3" runat="server">重要程度：</asp:label></FONT></TD>
														<TD class="item-text" style="HEIGHT: 21px" vAlign="top" height="21"><select id="LevelList" style="COLOR: #990000" name="Color" runat="server">
																<option style="COLOR: #000000; BACKGROUND-COLOR: red" value="1">&nbsp;&nbsp;高&nbsp;&nbsp;</option>
																<option style="COLOR: #000000; BACKGROUND-COLOR: gold" value="2">&nbsp;&nbsp;中&nbsp;&nbsp;</option>
																<option style="COLOR: #000000; BACKGROUND-COLOR: #00c000" value="3" selected>&nbsp;&nbsp;低&nbsp;&nbsp;</option>
															</select>
														</TD>
													</TR>
													<TR>
														<TD class="header-item" style="HEIGHT: 20px" vAlign="top" align="right" width="120"
															height="20"></TD>
														<TD class="item-text" style="HEIGHT: 20px" vAlign="top" height="20"><FONT face="新細明體"></FONT></TD>
													</TR>
													<TR>
														<TD class="header-item" style="HEIGHT: 20px" vAlign="top" align="right" width="120"
															height="20"><FONT face="新細明體">地點：</FONT></TD>
														<TD class="item-text" style="HEIGHT: 20px" vAlign="top" height="20"><asp:textbox id="txtLocation" runat="server"></asp:textbox></TD>
													</TR>
													<TR vAlign="top">
														<TD class="header-item" style="HEIGHT: 20px" vAlign="top" align="right" width="120"
															height="20">&nbsp;</TD>
														<TD class="item-text" style="HEIGHT: 20px" vAlign="top" height="20"><FONT face="新細明體"></FONT></TD>
													</TR>
													<TR>
														<TD class="header-item" style="HEIGHT: 19px" align="right" width="120"><FONT face="新細明體">時間設定：</FONT></TD>
														<TD class="item-text" style="HEIGHT: 19px"><FONT face="新細明體"><asp:radiobutton id="rbtnSigle" runat="server" GroupName="SingleCycle" Text="單一事項" Checked="True"
																	AutoPostBack="True"></asp:radiobutton><asp:radiobutton id="rbtnCycle" runat="server" GroupName="SingleCycle" Text="週期事項" AutoPostBack="True"></asp:radiobutton></FONT></TD>
													</TR>
													<TR>
														<TD class="header-item" style="HEIGHT: 19px" align="right" width="120"></TD>
														<TD class="item-text" style="HEIGHT: 19px" height="19"></TD>
													</TR>
													<TR>
														<TD class="header-item" align="right" width="120"><FONT face="新細明體"></FONT></TD>
														<TD class="item-text" vAlign="top">
															<P><asp:panel id="PanelSingle" runat="server"></P>
															<FIELDSET class="fieldset" style="WIDTH: 448px; HEIGHT: 80px"><LEGEND id="Legend2">事項時間</LEGEND>
																<DIV style="PADDING-RIGHT: 3px; PADDING-LEFT: 3px; PADDING-BOTTOM: 3px; PADDING-TOP: 3px">
																	<TABLE id="Table8" cellSpacing="0" cellPadding="0" width="100%" border="0">
																		<TR>
																			<TD style="HEIGHT: 21px" colSpan="2"><asp:label id="Label7" runat="server">日期起迄：</asp:label><cc1:datepicker id="txtSingleStartDate" runat="server" ControlCssClass="DatePicker" DateType="yyyy/mm/dd"
																					imgDirectory="/app/outerapp/ProjectManagement/datePicker/"></cc1:datepicker><FONT face="新細明體">&nbsp;∼<cc1:datepicker id="txtSingleEndDate" runat="server" ControlCssClass="DatePicker" DateType="yyyy/mm/dd"
																						imgDirectory="/app/outerapp/ProjectManagement/datePicker/"></cc1:datepicker></FONT></TD>
																		</TR>
																		<TR>
																			<TD colSpan="2"><asp:label id="Label5" runat="server">時間起迄：</asp:label><asp:listbox id="lbxSingleStartTime" runat="server" Rows="1">
																					<asp:ListItem Value="00:30">00:30</asp:ListItem>
																					<asp:ListItem Value="01:00">01:00</asp:ListItem>
																					<asp:ListItem Value="01:30">01:30</asp:ListItem>
																					<asp:ListItem Value="02:00">02:00</asp:ListItem>
																					<asp:ListItem Value="02:30">02:30</asp:ListItem>
																					<asp:ListItem Value="03:00">03:00</asp:ListItem>
																					<asp:ListItem Value="03:30">03:30</asp:ListItem>
																					<asp:ListItem Value="04:00">04:00</asp:ListItem>
																					<asp:ListItem Value="04:30">04:30</asp:ListItem>
																					<asp:ListItem Value="05:00">05:00</asp:ListItem>
																					<asp:ListItem Value="05:30">05:30</asp:ListItem>
																					<asp:ListItem Value="06:00">06:00</asp:ListItem>
																					<asp:ListItem Value="06:30">06:30</asp:ListItem>
																					<asp:ListItem Value="07:00">07:00</asp:ListItem>
																					<asp:ListItem Value="07:30">07:30</asp:ListItem>
																					<asp:ListItem Value="08:00" Selected="True">08:00</asp:ListItem>
																					<asp:ListItem Value="08:30">08:30</asp:ListItem>
																					<asp:ListItem Value="09:00">09:00</asp:ListItem>
																					<asp:ListItem Value="09:30">09:30</asp:ListItem>
																					<asp:ListItem Value="10:00">10:00</asp:ListItem>
																					<asp:ListItem Value="10:30">10:30</asp:ListItem>
																					<asp:ListItem Value="11:00">11:00</asp:ListItem>
																					<asp:ListItem Value="11:30">11:30</asp:ListItem>
																					<asp:ListItem Value="12:00">12:00</asp:ListItem>
																					<asp:ListItem Value="12:30">12:30</asp:ListItem>
																					<asp:ListItem Value="13:00">13:00</asp:ListItem>
																					<asp:ListItem Value="13:30">13:30</asp:ListItem>
																					<asp:ListItem Value="14:00">14:00</asp:ListItem>
																					<asp:ListItem Value="14:30">14:30</asp:ListItem>
																					<asp:ListItem Value="15:00">15:00</asp:ListItem>
																					<asp:ListItem Value="15:30">15:30</asp:ListItem>
																					<asp:ListItem Value="16:00">16:00</asp:ListItem>
																					<asp:ListItem Value="16:30">16:30</asp:ListItem>
																					<asp:ListItem Value="17:00">17:00</asp:ListItem>
																					<asp:ListItem Value="17:30">17:30</asp:ListItem>
																					<asp:ListItem Value="18:00">18:00</asp:ListItem>
																					<asp:ListItem Value="18:30">18:30</asp:ListItem>
																					<asp:ListItem Value="19:00">19:00</asp:ListItem>
																					<asp:ListItem Value="19:30">19:30</asp:ListItem>
																					<asp:ListItem Value="20:00">20:00</asp:ListItem>
																					<asp:ListItem Value="20:30">20:30</asp:ListItem>
																					<asp:ListItem Value="21:00">21:00</asp:ListItem>
																					<asp:ListItem Value="21:30">21:30</asp:ListItem>
																					<asp:ListItem Value="22:00">22:00</asp:ListItem>
																					<asp:ListItem Value="22:30">22:30</asp:ListItem>
																					<asp:ListItem Value="23:00">23:00</asp:ListItem>
																					<asp:ListItem Value="23:30">23:30</asp:ListItem>
																				</asp:listbox><FONT face="新細明體">∼</FONT><asp:listbox id="lbxSingleEndTime" runat="server" Rows="1">
																					<asp:ListItem Value="00:30">00:30</asp:ListItem>
																					<asp:ListItem Value="01:00">01:00</asp:ListItem>
																					<asp:ListItem Value="01:30">01:30</asp:ListItem>
																					<asp:ListItem Value="02:00">02:00</asp:ListItem>
																					<asp:ListItem Value="02:30">02:30</asp:ListItem>
																					<asp:ListItem Value="03:00">03:00</asp:ListItem>
																					<asp:ListItem Value="03:30">03:30</asp:ListItem>
																					<asp:ListItem Value="04:00">04:00</asp:ListItem>
																					<asp:ListItem Value="04:30">04:30</asp:ListItem>
																					<asp:ListItem Value="05:00">05:00</asp:ListItem>
																					<asp:ListItem Value="05:30">05:30</asp:ListItem>
																					<asp:ListItem Value="06:00">06:00</asp:ListItem>
																					<asp:ListItem Value="06:30">06:30</asp:ListItem>
																					<asp:ListItem Value="07:00">07:00</asp:ListItem>
																					<asp:ListItem Value="07:30">07:30</asp:ListItem>
																					<asp:ListItem Value="08:00" Selected="True">08:00</asp:ListItem>
																					<asp:ListItem Value="08:30">08:30</asp:ListItem>
																					<asp:ListItem Value="09:00">09:00</asp:ListItem>
																					<asp:ListItem Value="09:30">09:30</asp:ListItem>
																					<asp:ListItem Value="10:00">10:00</asp:ListItem>
																					<asp:ListItem Value="10:30">10:30</asp:ListItem>
																					<asp:ListItem Value="11:00">11:00</asp:ListItem>
																					<asp:ListItem Value="11:30">11:30</asp:ListItem>
																					<asp:ListItem Value="12:00">12:00</asp:ListItem>
																					<asp:ListItem Value="12:30">12:30</asp:ListItem>
																					<asp:ListItem Value="13:00">13:00</asp:ListItem>
																					<asp:ListItem Value="13:30">13:30</asp:ListItem>
																					<asp:ListItem Value="14:00">14:00</asp:ListItem>
																					<asp:ListItem Value="14:30">14:30</asp:ListItem>
																					<asp:ListItem Value="15:00">15:00</asp:ListItem>
																					<asp:ListItem Value="15:30">15:30</asp:ListItem>
																					<asp:ListItem Value="16:00">16:00</asp:ListItem>
																					<asp:ListItem Value="16:30">16:30</asp:ListItem>
																					<asp:ListItem Value="17:00">17:00</asp:ListItem>
																					<asp:ListItem Value="17:30">17:30</asp:ListItem>
																					<asp:ListItem Value="18:00">18:00</asp:ListItem>
																					<asp:ListItem Value="18:30">18:30</asp:ListItem>
																					<asp:ListItem Value="19:00">19:00</asp:ListItem>
																					<asp:ListItem Value="19:30">19:30</asp:ListItem>
																					<asp:ListItem Value="20:00">20:00</asp:ListItem>
																					<asp:ListItem Value="20:30">20:30</asp:ListItem>
																					<asp:ListItem Value="21:00">21:00</asp:ListItem>
																					<asp:ListItem Value="21:30">21:30</asp:ListItem>
																					<asp:ListItem Value="22:00">22:00</asp:ListItem>
																					<asp:ListItem Value="22:30">22:30</asp:ListItem>
																					<asp:ListItem Value="23:00">23:00</asp:ListItem>
																					<asp:ListItem Value="23:30">23:30</asp:ListItem>
																				</asp:listbox></TD>
																		</TR>
																	</TABLE>
																</DIV>
															</FIELDSET>
															</asp:panel></FONT><asp:panel id="PanelCycle" runat="server" Width="440px" Visible="False">
																<FIELDSET class="fieldset" style="WIDTH: 448px; HEIGHT: 80px"><LEGEND id="LEGEND1">週期時間</LEGEND>
																	<DIV style="PADDING-RIGHT: 3px; PADDING-LEFT: 3px; PADDING-BOTTOM: 3px; PADDING-TOP: 3px">
																		<TABLE id="Table6" cellSpacing="0" cellPadding="0" width="100%" border="0">
																			<TR>
																				<TD style="HEIGHT: 21px" colSpan="2">
																					<asp:label id="Label4" runat="server">日期起迄：</asp:label>
																					<cc1:datepicker id="txtCycleStartDate" runat="server" imgDirectory="/app/outerapp/ProjectManagement/datePicker/"
																						DateType="yyyy/mm/dd" ControlCssClass="DatePicker"></cc1:datepicker><FONT face="新細明體">∼
																						<cc1:datepicker id="txtCycleEndDate" runat="server" imgDirectory="/app/outerapp/ProjectManagement/datePicker/"
																							DateType="yyyy/mm/dd" ControlCssClass="DatePicker"></cc1:datepicker></FONT></TD>
																			</TR>
																			<TR>
																				<TD colSpan="2">
																					<asp:label id="Label6" runat="server">時間起訖：</asp:label>
																					<asp:listbox id="lbxCycleStartTime" runat="server" Rows="1">
																						<asp:ListItem Value="00:30">00:30</asp:ListItem>
																						<asp:ListItem Value="01:00">01:00</asp:ListItem>
																						<asp:ListItem Value="01:30">01:30</asp:ListItem>
																						<asp:ListItem Value="02:00">02:00</asp:ListItem>
																						<asp:ListItem Value="02:30">02:30</asp:ListItem>
																						<asp:ListItem Value="03:00">03:00</asp:ListItem>
																						<asp:ListItem Value="03:30">03:30</asp:ListItem>
																						<asp:ListItem Value="04:00">04:00</asp:ListItem>
																						<asp:ListItem Value="04:30">04:30</asp:ListItem>
																						<asp:ListItem Value="05:00">05:00</asp:ListItem>
																						<asp:ListItem Value="05:30">05:30</asp:ListItem>
																						<asp:ListItem Value="06:00">06:00</asp:ListItem>
																						<asp:ListItem Value="06:30">06:30</asp:ListItem>
																						<asp:ListItem Value="07:00">07:00</asp:ListItem>
																						<asp:ListItem Value="07:30">07:30</asp:ListItem>
																						<asp:ListItem Value="08:00" Selected="True">08:00</asp:ListItem>
																						<asp:ListItem Value="08:30">08:30</asp:ListItem>
																						<asp:ListItem Value="09:00">09:00</asp:ListItem>
																						<asp:ListItem Value="09:30">09:30</asp:ListItem>
																						<asp:ListItem Value="10:00">10:00</asp:ListItem>
																						<asp:ListItem Value="10:30">10:30</asp:ListItem>
																						<asp:ListItem Value="11:00">11:00</asp:ListItem>
																						<asp:ListItem Value="11:30">11:30</asp:ListItem>
																						<asp:ListItem Value="12:00">12:00</asp:ListItem>
																						<asp:ListItem Value="12:30">12:30</asp:ListItem>
																						<asp:ListItem Value="13:00">13:00</asp:ListItem>
																						<asp:ListItem Value="13:30">13:30</asp:ListItem>
																						<asp:ListItem Value="14:00">14:00</asp:ListItem>
																						<asp:ListItem Value="14:30">14:30</asp:ListItem>
																						<asp:ListItem Value="15:00">15:00</asp:ListItem>
																						<asp:ListItem Value="15:30">15:30</asp:ListItem>
																						<asp:ListItem Value="16:00">16:00</asp:ListItem>
																						<asp:ListItem Value="16:30">16:30</asp:ListItem>
																						<asp:ListItem Value="17:00">17:00</asp:ListItem>
																						<asp:ListItem Value="17:30">17:30</asp:ListItem>
																						<asp:ListItem Value="18:00">18:00</asp:ListItem>
																						<asp:ListItem Value="18:30">18:30</asp:ListItem>
																						<asp:ListItem Value="19:00">19:00</asp:ListItem>
																						<asp:ListItem Value="19:30">19:30</asp:ListItem>
																						<asp:ListItem Value="20:00">20:00</asp:ListItem>
																						<asp:ListItem Value="20:30">20:30</asp:ListItem>
																						<asp:ListItem Value="21:00">21:00</asp:ListItem>
																						<asp:ListItem Value="21:30">21:30</asp:ListItem>
																						<asp:ListItem Value="22:00">22:00</asp:ListItem>
																						<asp:ListItem Value="22:30">22:30</asp:ListItem>
																						<asp:ListItem Value="23:00">23:00</asp:ListItem>
																						<asp:ListItem Value="23:30">23:30</asp:ListItem>
																					</asp:listbox><FONT face="新細明體">∼</FONT>
																					<asp:listbox id="lbxCycleEndTime" runat="server" Rows="1">
																						<asp:ListItem Value="00:30">00:30</asp:ListItem>
																						<asp:ListItem Value="01:00">01:00</asp:ListItem>
																						<asp:ListItem Value="01:30">01:30</asp:ListItem>
																						<asp:ListItem Value="02:00">02:00</asp:ListItem>
																						<asp:ListItem Value="02:30">02:30</asp:ListItem>
																						<asp:ListItem Value="03:00">03:00</asp:ListItem>
																						<asp:ListItem Value="03:30">03:30</asp:ListItem>
																						<asp:ListItem Value="04:00">04:00</asp:ListItem>
																						<asp:ListItem Value="04:30">04:30</asp:ListItem>
																						<asp:ListItem Value="05:00">05:00</asp:ListItem>
																						<asp:ListItem Value="05:30">05:30</asp:ListItem>
																						<asp:ListItem Value="06:00">06:00</asp:ListItem>
																						<asp:ListItem Value="06:30">06:30</asp:ListItem>
																						<asp:ListItem Value="07:00">07:00</asp:ListItem>
																						<asp:ListItem Value="07:30">07:30</asp:ListItem>
																						<asp:ListItem Value="08:00" Selected="True">08:00</asp:ListItem>
																						<asp:ListItem Value="08:30">08:30</asp:ListItem>
																						<asp:ListItem Value="09:00">09:00</asp:ListItem>
																						<asp:ListItem Value="09:30">09:30</asp:ListItem>
																						<asp:ListItem Value="10:00">10:00</asp:ListItem>
																						<asp:ListItem Value="10:30">10:30</asp:ListItem>
																						<asp:ListItem Value="11:00">11:00</asp:ListItem>
																						<asp:ListItem Value="11:30">11:30</asp:ListItem>
																						<asp:ListItem Value="12:00">12:00</asp:ListItem>
																						<asp:ListItem Value="12:30">12:30</asp:ListItem>
																						<asp:ListItem Value="13:00">13:00</asp:ListItem>
																						<asp:ListItem Value="13:30">13:30</asp:ListItem>
																						<asp:ListItem Value="14:00">14:00</asp:ListItem>
																						<asp:ListItem Value="14:30">14:30</asp:ListItem>
																						<asp:ListItem Value="15:00">15:00</asp:ListItem>
																						<asp:ListItem Value="15:30">15:30</asp:ListItem>
																						<asp:ListItem Value="16:00">16:00</asp:ListItem>
																						<asp:ListItem Value="16:30">16:30</asp:ListItem>
																						<asp:ListItem Value="17:00">17:00</asp:ListItem>
																						<asp:ListItem Value="17:30">17:30</asp:ListItem>
																						<asp:ListItem Value="18:00">18:00</asp:ListItem>
																						<asp:ListItem Value="18:30">18:30</asp:ListItem>
																						<asp:ListItem Value="19:00">19:00</asp:ListItem>
																						<asp:ListItem Value="19:30">19:30</asp:ListItem>
																						<asp:ListItem Value="20:00">20:00</asp:ListItem>
																						<asp:ListItem Value="20:30">20:30</asp:ListItem>
																						<asp:ListItem Value="21:00">21:00</asp:ListItem>
																						<asp:ListItem Value="21:30">21:30</asp:ListItem>
																						<asp:ListItem Value="22:00">22:00</asp:ListItem>
																						<asp:ListItem Value="22:30">22:30</asp:ListItem>
																						<asp:ListItem Value="23:00">23:00</asp:ListItem>
																						<asp:ListItem Value="23:30">23:30</asp:ListItem>
																					</asp:listbox></TD>
																			</TR>
																		</TABLE>
																	</DIV>
																</FIELDSET>
																<P></P>
																<FIELDSET class="fieldset" style="HEIGHT: 144px"><LEGEND>週期模式</LEGEND>
																	<DIV style="PADDING-RIGHT: 3px; PADDING-LEFT: 3px; PADDING-BOTTOM: 3px; PADDING-TOP: 3px">
																		<TABLE id="Table7" cellSpacing="0" cellPadding="0" width="100%" border="0">
																			<TR>
																				<TD style="WIDTH: 44px" vAlign="top"><BR>
																					<asp:radiobutton id="rbtnCycleType1" runat="server" AutoPostBack="True" Checked="True" Text="每日"
																						GroupName="CycleMode"></asp:radiobutton><BR>
																					<asp:radiobutton id="rbtnCycleType2" runat="server" AutoPostBack="True" Text="每週" GroupName="CycleMode"></asp:radiobutton><BR>
																					<asp:radiobutton id="rbtnCycleType3" runat="server" AutoPostBack="True" Text="每月" GroupName="CycleMode"></asp:radiobutton></TD>
																				<TD><FONT face="新細明體">&nbsp; &nbsp;&nbsp;</FONT>
																				</TD>
																				<TD><BR>
																					<FIELDSET class="fieldset" style="WIDTH: 376px; HEIGHT: 100px">
																						<DIV style="PADDING-RIGHT: 3px; PADDING-LEFT: 3px; PADDING-BOTTOM: 3px; PADDING-TOP: 3px">
																							<asp:panel id="PanelCycleDay" runat="server" Width="336px">
																								<P><FONT face="新細明體"></FONT>&nbsp;</P>
																								<P>
																									<asp:RadioButton id="rbtnCycleDayType1" runat="server" Checked="True" Text="每一天" GroupName="CycleDay"></asp:RadioButton>
																									<asp:RadioButton id="rbtnCycleDayType2" runat="server" Text="每個工作天" GroupName="CycleDay"></asp:RadioButton></P>
																							</asp:panel></DIV>
																						<asp:panel id="PanelCycleWeek" runat="server" Width="336px" Visible="False">每 
<asp:ListBox id="lboxWeekInterval" runat="server" Rows="1">
																								<asp:ListItem Value="1">1</asp:ListItem>
																								<asp:ListItem Value="2">2</asp:ListItem>
																								<asp:ListItem Value="3">3</asp:ListItem>
																								<asp:ListItem Value="4">4</asp:ListItem>
																								<asp:ListItem Value="5">5</asp:ListItem>
																							</asp:ListBox>週的 <BR><BR>
<asp:CheckBox id="cboxWeekDay1" runat="server" Text="星期一"></asp:CheckBox>
<asp:CheckBox id="cboxWeekDay2" runat="server" Text="星期二"></asp:CheckBox>
<asp:CheckBox id="cboxWeekDay3" runat="server" Text="星期三"></asp:CheckBox>
<asp:CheckBox id="cboxWeekDay4" runat="server" Text="星期四"></asp:CheckBox>
<asp:CheckBox id="cboxWeekDay5" runat="server" Text="星期五"></asp:CheckBox><BR>
<asp:CheckBox id="cboxWeekDay6" runat="server" Text="星期六"></asp:CheckBox>
<asp:CheckBox id="cboxWeekDay7" runat="server" Text="星期日"></asp:CheckBox></asp:panel>
																						<asp:panel id="PanelCycleMonth" runat="server" Width="336px" Visible="False">
<asp:RadioButton id="rbtnCycleMonthType1" runat="server" Checked="True" GroupName="CycleMonth"></asp:RadioButton>每月的 
<asp:ListBox id="lboxMonthDay" runat="server" Rows="1"></asp:ListBox>日<BR><BR>
<asp:RadioButton id="rbtnCycleMonthType2" runat="server" GroupName="CycleMonth"></asp:RadioButton>每月的第 
<asp:ListBox id="lboxMonthWeekInterval" runat="server" Rows="1">
																								<asp:ListItem Value="1">第一個</asp:ListItem>
																								<asp:ListItem Value="2">第二個</asp:ListItem>
																								<asp:ListItem Value="3">第三個</asp:ListItem>
																								<asp:ListItem Value="4">第四個</asp:ListItem>
																								<asp:ListItem Value="5">最後一個</asp:ListItem>
																							</asp:ListBox>
<asp:ListBox id="lboxMonthWeekDay" runat="server" Rows="1"></asp:ListBox></asp:panel></FIELDSET>
																				</TD>
																			</TR>
																		</TABLE>
																	</DIV>
																</FIELDSET>
															</asp:panel></TD>
													</TR>
													<TR>
														<TD class="header-item" vAlign="top" align="right" width="120" height="20"><FONT face="新細明體"></FONT></TD>
														<TD class="item-text" vAlign="top" height="20"><FONT face="新細明體"></FONT></TD>
													</TR>
													<TR vAlign="top">
														<TD class="header-item" vAlign="top" align="right" width="120" height="20">&nbsp;</TD>
														<TD class="item-text" vAlign="top" height="20"><FONT face="新細明體"></FONT></TD>
													</TR>
													<TR>
														<TD class="header-item" vAlign="top" align="right" width="120" height="20">&nbsp;
															<BR>
															<asp:label id="Label1" runat="server">備註：</asp:label></TD>
														<TD class="item-text" vAlign="top" height="20"><asp:textbox id="txtDescription" runat="server" Width="400px" Rows="5" TextMode="MultiLine"></asp:textbox></TD>
													</TR>
													<TR vAlign="top">
														<TD class="header-item" vAlign="top" align="right" width="120" height="20">&nbsp;</TD>
														<TD class="item-text" vAlign="top" height="20"><FONT face="新細明體"></FONT></TD>
													</TR>
													<TR>
														<TD class="header-item" vAlign="top" align="right" width="120" height="20"><BR>
														</TD>
														<TD class="item-text" vAlign="top" height="20"></TD>
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
		</FONT></TR></TBODY>
		<DIV></DIV>
		<FIELDSET></FIELDSET>
		</TR></TBODY></TABLE></TR></TBODY></TABLE></TR></TBODY></TABLE></FORM>
	</body>
</HTML>
