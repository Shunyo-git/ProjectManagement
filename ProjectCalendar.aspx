<%@ Page Language="vb" AutoEventWireup="false" Codebehind="ProjectCalendar.aspx.vb" Inherits="HsinYi.EIP.ProjectManagement.Web.ProjectCalendar"%>
<%@ Register TagPrefix="uc1" TagName="projectTabs" Src="inc/projectTabs.ascx" %>
<%@ Register TagPrefix="uc1" TagName="myProjectTabs" Src="inc/mainTitleTabs.ascx" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>專案行事曆</title>
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
						<TD vAlign="top" height="15"><IMG height="15" src="images/spacer.gif" width="15"></TD>
					</TR>
				</TABLE>
				<TABLE id="Table2" cellSpacing="0" cellPadding="0" width="100%" border="0">
					<TR>
						<TD width="8"><IMG height="8" src="images/spacer.gif" width="8"></TD>
						<TD vAlign="top"><uc1:projecttabs id="ProjectTabs1" runat="server"></uc1:projecttabs>
							<TABLE class="admin-tan-border" id="Table3" height="530" cellSpacing="20" cellPadding="0"
								width="100%" border="0">
								<TR>
									<TD height="20"><asp:panel id="PanelModify" runat="server">
											<TABLE id="Table6" cellSpacing="0" cellPadding="0" width="100%" border="0">
												<TR>
													<TD style="HEIGHT: 14px"><FONT face="新細明體"></FONT></TD>
													<TD style="HEIGHT: 14px" noWrap align="right"><IMG alt="" src="images/icon01.gif" align="absBottom">&nbsp;
														<asp:LinkButton id="lbtnNew" runat="server">新增事項</asp:LinkButton>&nbsp;
													</TD>
												</TR>
												<TR>
													<TD noWrap align="right" colSpan="2"><IMG alt="" src="images/l_ba_bottom.GIF">
													</TD>
												</TR>
											</TABLE>
										</asp:panel></TD>
								</TR>
								<TR vAlign="top" height="*">
									<TD align="center"><asp:calendar id="PjCalendar" runat="server" SelectedDayStyle-BackColor="Navy" SelectedDate="1/1/0001"
											DayStyle-Width="14%" DayStyle-Height="50px" DayStyle-VerticalAlign="Top" TitleStyle-Font-Bold="true" TitleStyle-Font-Size="12px"
											TitleStyle-BackColor="Gainsboro" VisibleDate="2000-09-16" Width="90%" Font-Size="X-Small" Font-Name="Verdana"
											BorderWidth="1" ShowGridLines="true" PrevMonthText="<img src='images/monthleft.gif' border=0>" NextMonthText="<img src='images/monthright.gif' border=0>"
											ForeColor="Black" Font-Names="Verdana">
											<DayStyle Font-Bold="True" HorizontalAlign="Left" Height="80px" Width="14%" VerticalAlign="Top"></DayStyle>
											<SelectedDayStyle BackColor="LightSkyBlue"></SelectedDayStyle>
											<TitleStyle Font-Size="12px" Font-Bold="True" BackColor="Gainsboro"></TitleStyle>
											<WeekendDayStyle BackColor="#FFFFC0"></WeekendDayStyle>
											<OtherMonthDayStyle ForeColor="Gray"></OtherMonthDayStyle>
										</asp:calendar></TD>
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
