<%@ Page Language="vb" AutoEventWireup="false" Codebehind="CalendarDetail.aspx.vb" Inherits="HsinYi.EIP.ProjectManagement.Web.CalendarDetail"%>
<%@ Register TagPrefix="uc1" TagName="projectTabs" Src="inc/projectTabs.ascx" %>
<%@ Register TagPrefix="uc1" TagName="myProjectTabs" Src="inc/mainTitleTabs.ascx" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>��ƾ䤺�e</title>
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
									<TR>
										<TD height="20"><FONT face="�s�ө���">
												<TABLE id="Table6" cellSpacing="0" cellPadding="0" width="100%" border="0">
													<TR>
														<TD></TD>
														<TD noWrap align="right"><asp:panel id="PanelModify" runat="server" Visible="False"><IMG alt="" src="images/icon01.gif" align="absBottom"> 
<asp:LinkButton id="lbtnEdit" runat="server">�ק�</asp:LinkButton>&nbsp;<IMG alt="" src="images/icon01.gif" align="absBottom"> 
<asp:LinkButton id="lbtnDelete" runat="server">�R���ƶ�</asp:LinkButton><IMG alt="" src="images/icon01.gif" align="absBottom"> 
<asp:LinkButton id="lbtnDeltetCycle" runat="server">�R����Ӷg��</asp:LinkButton></asp:panel></TD>
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
										<TD><FONT face="�s�ө���">
												<TABLE id="Table4" cellSpacing="0" cellPadding="0" width="100%" border="0">
													<TBODY>
														<TR>
															<TD class="header-item" style="HEIGHT: 21px" align="right" width="120" height="21">�ƶ��G&nbsp;
															</TD>
															<TD class="item-text" style="HEIGHT: 21px" height="21"><FONT face="�s�ө���"><asp:label id="lblSubject" runat="server">�ƶ��G</asp:label></FONT></TD>
														</TR>
														<TR>
															<TD class="header-item" style="HEIGHT: 20px" vAlign="top" align="right" width="120"
																height="20"></TD>
															<TD class="item-text" style="HEIGHT: 20px" vAlign="top" height="20"></TD>
														</TR>
														<TR>
															<TD class="header-item" style="HEIGHT: 13px" vAlign="top" align="right" width="120"
																height="13"><FONT face="�s�ө���">�����G</FONT></TD>
															<TD class="item-text" style="HEIGHT: 13px" vAlign="top" height="13"><FONT face="�s�ө���"><asp:label id="lblAppointType" runat="server">�����G</asp:label></FONT></TD>
														</TR>
														<TR>
															<TD class="header-item" style="HEIGHT: 20px" vAlign="top" align="right" width="120"
																height="20"></TD>
															<TD class="item-text" style="HEIGHT: 20px" vAlign="top" height="20"></TD>
														</TR>
														<TR>
															<TD class="header-item" style="HEIGHT: 19px" vAlign="top" align="right" width="120"
																height="19"><FONT face="�s�ө���">���n�{�סG</FONT></TD>
															<TD class="item-text" style="HEIGHT: 19px" vAlign="top" height="19"><asp:label id="lblAppointLevel" runat="server">���n�{�סG</asp:label></TD>
														</TR>
														<TR>
															<TD class="header-item" style="HEIGHT: 20px" vAlign="top" align="right" width="120"
																height="20"></TD>
															<TD class="item-text" style="HEIGHT: 20px" vAlign="top" height="20"><FONT face="�s�ө���"></FONT></TD>
														</TR>
														<TR>
															<TD class="header-item" style="HEIGHT: 20px" vAlign="top" align="right" width="120"
																height="20"><FONT face="�s�ө���">�a�I�G</FONT></TD>
															<TD class="item-text" style="HEIGHT: 20px" vAlign="top" height="20"><asp:label id="lblLocation" runat="server"></asp:label></TD>
														</TR>
														<TR vAlign="top">
															<TD class="header-item" style="HEIGHT: 20px" vAlign="top" align="right" width="120"
																height="20">&nbsp;</TD>
															<TD class="item-text" style="HEIGHT: 20px" vAlign="top" height="20"><FONT face="�s�ө���"></FONT></TD>
														</TR>
														<TR>
															<TD class="header-item" style="HEIGHT: 19px" align="right" width="120"><FONT face="�s�ө���">�ɶ������G</FONT></TD>
															<TD class="item-text" style="HEIGHT: 19px"><FONT face="�s�ө���"><asp:label id="lblSingle" runat="server">��@�ƶ�</asp:label><asp:label id="lblCycle" runat="server"><img src="images/icon_cycle.gif">�g���ƶ�</asp:label></FONT></TD>
														</TR>
														<TR>
															<TD class="header-item" style="HEIGHT: 20px" align="right" width="120"></TD>
															<TD class="item-text" style="HEIGHT: 20px" height="14"></TD>
														</TR>
														<TR>
															<TD class="header-item" align="right" width="120"><FONT face="�s�ө���"></FONT></TD>
															<TD class="item-text" vAlign="top">
																<P><asp:panel id="PanelSingle" runat="server"></P>
																<FIELDSET class="fieldset" style="WIDTH: 448px; HEIGHT: 80px"><LEGEND id="Legend2">�ƶ��ɶ�</LEGEND>
																	<DIV style="PADDING-RIGHT: 3px; PADDING-LEFT: 3px; PADDING-BOTTOM: 3px; PADDING-TOP: 3px">
																		<TABLE class="item-text" id="Table8" cellSpacing="0" cellPadding="0" width="100%" border="0">
																			<TR>
																				<TD style="HEIGHT: 21px" colSpan="2"><FONT class="item-text" face="�s�ө���"><FONT class="item-text" face="�s�ө���">�ɶ�</FONT>�_�W�G
																						<asp:label id="lblSingleDateStart" runat="server"></asp:label>��
																						<asp:label id="lblSingleDateEnd" runat="server"></asp:label></FONT></TD>
																			</TR>
																		</TABLE>
																	</DIV>
																</FIELDSET>
												</asp:panel></FONT><asp:panel id="PanelCycle" runat="server" Visible="False" Width="440px">
												<FIELDSET class="fieldset" style="WIDTH: 448px; HEIGHT: 80px"><LEGEND id="LEGEND1">�g���ɶ�</LEGEND>
													<DIV style="PADDING-RIGHT: 3px; PADDING-LEFT: 3px; PADDING-BOTTOM: 3px; PADDING-TOP: 3px">
														<TABLE id="Table5" cellSpacing="0" cellPadding="0" width="100%" border="0">
															<TR>
																<TD style="HEIGHT: 21px" colSpan="2"><FONT class="item-text" face="�s�ө���">����_�W�G
																		<asp:label id="lblCycleDateStart" runat="server"></asp:label>��
																		<asp:label id="lblCycleDateEnd" runat="server"></asp:label></FONT></TD>
															</TR>
															<TR>
																<TD colSpan="2"><FONT class="item-text" face="�s�ө���">�ɶ��_���G
																		<asp:label id="lblCycleTimeStart" runat="server"></asp:label>��
																		<asp:label id="lblCycleTimeEnd" runat="server"></asp:label></FONT></TD>
															</TR>
														</TABLE>
													</DIV>
												</FIELDSET>
												<P></P>
												<FIELDSET class="fieldset"><LEGEND>�g���Ҧ�</LEGEND>
													<DIV style="PADDING-RIGHT: 3px; PADDING-LEFT: 3px; PADDING-BOTTOM: 3px; PADDING-TOP: 3px">
														<TABLE class="item-text" id="Table7" cellSpacing="0" cellPadding="0" width="100%" border="0">
															<TR>
																<TD class="item-text" style="WIDTH: 44px" vAlign="top" colSpan="3">
																	<asp:panel id="PanelCycleDay" runat="server" Width="336px">
																		<P>
																			<asp:Label id="lblCycleDayType1" runat="server">�C�@��</asp:Label>
																			<asp:Label id="lblCycleDayType2" runat="server">�C�@�u�@��</asp:Label></P>
																	</asp:panel>
																	<asp:panel id="PanelCycleWeek" runat="server" Visible="False" Width="336px">�C 
<asp:Label id="lblCycleWeelInterval" runat="server">�@</asp:Label>�g�� 
                        <BR><BR>
<asp:Label id="lblWeekDay1" runat="server" Visible="False">�P���@ .</asp:Label>
<asp:Label id="lblWeekDay2" runat="server" Visible="False">�P���G .</asp:Label>
<asp:Label id="lblWeekDay3" runat="server" Visible="False">�P���T .</asp:Label>
<asp:Label id="lblWeekDay4" runat="server" Visible="False">�P���| . </asp:Label>
<asp:Label id="lblWeekDay5" runat="server" Visible="False">�P���� .</asp:Label>
<asp:Label id="lblWeekDay6" runat="server" Visible="False">�P���� .</asp:Label>
<asp:Label id="lblWeekDay7" runat="server" Visible="False">�P���� .</asp:Label></asp:panel>
																	<asp:panel id="PanelCycleMonth" runat="server" Visible="False" Width="336px">
																		<P>�C�몺&nbsp;
																			<asp:Label id="lblCycleMonthDay" runat="server">1 ��</asp:Label>
																			<asp:Label id="lblCycleMonthInterval" runat="server">�Ĥ@��</asp:Label>
																			<asp:Label id="lblCyclMonthWeekDay" runat="server">�P���@</asp:Label></P>
																	</asp:panel></TD>
															</TR>
														</TABLE>
													</DIV>
												</FIELDSET>
											</asp:panel></TD>
									</TR>
									<TR>
										<TD class="header-item" vAlign="top" align="right" width="120" height="20"><FONT face="�s�ө���"></FONT></TD>
										<TD class="item-text" vAlign="top" height="20"><FONT face="�s�ө���"></FONT></TD>
									</TR>
									<TR vAlign="top">
										<TD class="header-item" vAlign="top" align="right" width="120" height="20">�Ƶ��G&nbsp;</TD>
										<TD class="item-text" vAlign="top" height="20"><FONT face="�s�ө���"><asp:label id="lblDescription" runat="server"></asp:label></FONT></TD>
									</TR>
									<TR>
										<TD class="header-item" vAlign="top" align="right" width="120" height="20">&nbsp;
											<BR>
										</TD>
										<TD class="item-text" vAlign="top" height="20"></TD>
									</TR>
									<TR vAlign="top">
										<TD class="header-item" vAlign="top" align="right" width="120" height="20">&nbsp;</TD>
										<TD class="item-text" vAlign="top" height="20"><FONT face="�s�ө���"></FONT></TD>
									</TR>
									<TR>
										<TD class="header-item" vAlign="top" align="right" width="120" height="20"><BR>
										</TD>
										<TD class="item-text" vAlign="top" height="20"></TD>
									</TR>
								</TABLE>
		</FONT></TD></TR>
		<TR vAlign="top">
			<TD colSpan="3" height="12"></TD>
		</TR>
		</TBODY></TABLE></TD>
		<TD width="11"><IMG height="11" src="images/spacer.gif" width="11"></TD>
		</TR>
		<TR>
			<TD vAlign="top" colSpan="5" height="15"><IMG height="15" src="images/spacer.gif" width="15"></TD>
		</TR>
		</TBODY></TABLE></FORM></FONT>
	</body>
</HTML>
