<%@ Control Language="vb" AutoEventWireup="false" Codebehind="myCalendar.ascx.vb" Inherits="HsinYi.EIP.ProjectManagement.Web.myCalendar" TargetSchema="http://schemas.microsoft.com/intellisense/ie5" %>
<FONT face="新細明體">
	<TABLE id="Table1" cellSpacing="0" cellPadding="0" width="160" border="0">
		<TR>
			<TD class="item-text"><asp:calendar id="PjCalendar" Font-Names="Verdana" ForeColor="Black" NextMonthText="<img src='images/monthright.gif' border=0>"
					PrevMonthText="<img src='images/monthleft.gif' border=0>" ShowGridLines="true" BorderWidth="1" Font-Name="Verdana"
					Font-Size="X-Small" VisibleDate="2000-09-16" TitleStyle-BackColor="Gainsboro" TitleStyle-Font-Size="12px" TitleStyle-Font-Bold="true"
					DayStyle-VerticalAlign="Top" DayStyle-Width="14%" SelectedDate="1/1/0001" SelectedDayStyle-BackColor="Navy"
					runat="server" Width="100%">
					<TodayDayStyle Font-Bold="True" ForeColor="White" BackColor="#990000"></TodayDayStyle>
					<DayStyle BorderWidth="2px" ForeColor="#666666" BorderStyle="Solid" BorderColor="White" Width="14%"
						VerticalAlign="Top" BackColor="#EAEAEA"></DayStyle>
					<DayHeaderStyle ForeColor="#649CBA"></DayHeaderStyle>
					<SelectedDayStyle BackColor="LightSkyBlue"></SelectedDayStyle>
					<TitleStyle Font-Size="12px" Font-Bold="True" BackColor="Gainsboro"></TitleStyle>
					<WeekendDayStyle ForeColor="White" BackColor="#BBBBBB"></WeekendDayStyle>
					<OtherMonthDayStyle ForeColor="#666666" BackColor="White"></OtherMonthDayStyle>
				</asp:calendar><BR>
				<asp:datagrid id="ProjectsGrid" runat="server" Width="100%" PageSize="5" BorderColor="White" BorderStyle="None"
					CellPadding="2" AutoGenerateColumns="False">
					<HeaderStyle CssClass="grid-header"></HeaderStyle>
					<Columns>
						<asp:TemplateColumn HeaderText="本日專案事項">
							<HeaderStyle HorizontalAlign="Left" Width="5%" CssClass="grid-header"></HeaderStyle>
							<ItemTemplate>
								&nbsp;<%# WriteAppoint(DataBinder.Eval(Container, "DataItem.ProjectID")) %>
							</ItemTemplate>
						</asp:TemplateColumn>
					</Columns>
					<PagerStyle HorizontalAlign="Right" PageButtonCount="5"></PagerStyle>
				</asp:datagrid></TD>
		</TR>
	</TABLE>
</FONT>
