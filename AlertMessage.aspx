<%@ Page Language="vb" AutoEventWireup="false" Codebehind="AlertMessage.aspx.vb" Inherits="AlertMessage"%>
<%@ Register TagPrefix="uc1" TagName="myProjectTabs" Src="inc/mainTitleTabs.ascx" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>訊息提示</title>
		<meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
		<meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
		<meta content="JavaScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
		<LINK href="css/Styles.css" type="text/css" rel="stylesheet">
		<script language="javascript">
		function goBack(){
			window.history.back();
			
		}
		</script>
	</HEAD>
	<body>
		<FONT face="新細明體">
			<FORM id="ProjectList" method="post" runat="server">
				<TABLE id="Table1" cellSpacing="0" cellPadding="0" width="100%" border="0">
					<TR>
						<TD vAlign="bottom" bgColor="buttonface" height="46"><FONT face="新細明體"><uc1:myprojecttabs id="MyProjectTabs1" runat="server"></uc1:myprojecttabs></FONT></TD>
					</TR>
					<TR>
						<TD vAlign="top" height="15"><IMG height="15" src="images/spacer.gif" width="15">&nbsp;
						</TD>
					</TR>
				</TABLE>
				<table style="HEIGHT: 156px" cellSpacing="0" cellPadding="0" width="90%" align="center"
					border="0">
					<tr>
						<td class="tab-table" style="HEIGHT: 153px" vAlign="top" height="153">
							<p><IMG height="15" src="images/spacer.gif" width="15"></p>
							<p align="left">
								<asp:label id="lblMsg" runat="server">錯誤訊息..</asp:label>
								<br>
							</p>
							<p><IMG height="15" src="images/spacer.gif" width="15"></p>
							<p align="center">
								<INPUT onclick="goBack();" type="button" value="回上一頁">
							</p>
						</td>
					</tr>
				</table>
				</TABLE>
			</FORM>
		</FONT>
	</body>
</HTML>
