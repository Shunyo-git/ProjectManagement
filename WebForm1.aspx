<%@ Page Language="vb" AutoEventWireup="false" Codebehind="WebForm1.aspx.vb" Inherits="WebForm1"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
	<head>
		<title>WebForm1</title>
		<meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
		<meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
		<meta name="vs_defaultClientScript" content="JavaScript">
		<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
	</head>
	<body MS_POSITIONING="GridLayout">
		<form id="Form1" method="post" runat="server">
			<h3>DataGrid Custom Paging Example</h3>
			<asp:DataGrid id="ItemsGrid" runat="server" BorderColor="black" BorderWidth="1" CellPadding="3"
				AllowPaging="true" AllowCustomPaging="true" AutoGenerateColumns="false" OnPageIndexChanged="Grid_Change">
				<PagerStyle NextPageText="Forward" PrevPageText="Back" Position="Bottom" PageButtonCount="5"
					BackColor="#00aaaa"></PagerStyle>
				<AlternatingItemStyle BackColor="yellow"></AlternatingItemStyle>
				<HeaderStyle BackColor="#00aaaa"></HeaderStyle>
				<Columns>
					<asp:BoundColumn HeaderText="Number" DataField="IntegerValue" />
					<asp:BoundColumn HeaderText="Item" DataField="StringValue" />
					<asp:BoundColumn HeaderText="Price" DataField="CurrencyValue" DataFormatString="{0:c}">
						<ItemStyle HorizontalAlign="right"></ItemStyle>
					</asp:BoundColumn>
				</Columns>
			</asp:DataGrid>
			<br>
			<asp:CheckBox id="CheckBox1" Text="Show page navigation" AutoPostBack="true" runat="server" />
		</form>
	</body>
</html>
