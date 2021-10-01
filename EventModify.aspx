<%@ Register TagPrefix="uc1" TagName="projectTabs" Src="inc/projectTabs.ascx" %>
<%@ Register TagPrefix="uc1" TagName="myProjectTabs" Src="inc/mainTitleTabs.ascx" %>
<%@ Register TagPrefix="cc1" Namespace="DatePicker.iX.Controls" Assembly="DatePicker" %>
<%@ Page Language="vb" AutoEventWireup="false" Codebehind="EventModify.aspx.vb" Inherits="HsinYi.EIP.ProjectManagement.Web.EventModify"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>�s��S��ƥ�</title>
		<meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
		<meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
		<meta content="JavaScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
		<LINK href="css/Styles.css" type="text/css" rel="stylesheet">
		<script language="JavaScript" src="script/Global.js" type="text/JavaScript"></script>
		<script language="javascript">
		function toClear(e) {
			//�R���ɮ׫�,�M���ɮ���ܬ������
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
									<TD style="HEIGHT: 48px" height="48"><FONT face="�s�ө���"><asp:panel id="PanelModify" runat="server">
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
											</asp:panel></FONT></TD>
								</TR>
								<TR vAlign="top" height="*">
									<TD><FONT face="�s�ө���">
											<TABLE id="Table4" cellSpacing="0" cellPadding="0" width="100%" border="0">
												<TR>
													<TD class="header-item" align="right" width="120" height="20">&nbsp;
														<asp:label id="itnTopic" runat="server">�D�D�G</asp:label></TD>
													<TD class="item-text" height="20"><FONT face="�s�ө���"><asp:textbox id="txtTopic" runat="server" Width="350px"></asp:textbox><asp:requiredfieldvalidator id="ProjectNameRequiredfieldvalidator" runat="server" ErrorMessage="�D�D���o�ť�." ControlToValidate="txtTopic"
																Display="Dynamic"></asp:requiredfieldvalidator></FONT></TD>
												</TR>
												<TR vAlign="top">
													<TD class="header-item" vAlign="top" align="right" width="120" height="20" style="HEIGHT: 20px">&nbsp;</TD>
													<TD class="item-text" vAlign="top" height="20" style="HEIGHT: 20px"><FONT face="�s�ө���"></FONT></TD>
												</TR>
												<TR>
													<TD class="header-item" align="right" width="120" height="20">&nbsp;
														<asp:label id="itnDate" runat="server">�ɶ��G</asp:label></TD>
													<TD class="item-text" height="20"><FONT face="�s�ө���"><cc1:datepicker id="txtEventDate" runat="server" ControlCssClass="DatePicker" DateType="yyyy/mm/dd"
																imgDirectory="/app/outerapp/ProjectManagement/datePicker/"></cc1:datepicker></FONT></TD>
												</TR>
												<TR vAlign="top">
													<TD class="header-item" vAlign="top" align="right" width="120" height="20">&nbsp;</TD>
													<TD class="item-text" vAlign="top" height="20"><FONT face="�s�ө���"></FONT></TD>
												</TR>
												<TR>
													<TD class="header-item" vAlign="top" align="right" width="120" height="20">&nbsp;
														<br>
														<asp:label id="Label1" runat="server">�ƥ󻡩��G</asp:label></TD>
													<TD class="item-text" vAlign="top" height="20"><asp:textbox id="txtReason" runat="server" Width="400px" TextMode="MultiLine" Rows="5"></asp:textbox></TD>
												</TR>
												<TR vAlign="top">
													<TD class="header-item" vAlign="top" align="right" width="120" height="20"></TD>
													<TD class="item-text" vAlign="top" height="20"><FONT face="�s�ө���"></FONT></TD>
												</TR>
												<TR>
													<TD class="header-item" vAlign="top" align="right" width="120" height="20"><br>
														<asp:label id="passwordlbl" runat="server">�B�z��k�G</asp:label></TD>
													<TD class="item-text" vAlign="top" height="20"><FONT face="�s�ө���"><asp:textbox id="txtAnswer" runat="server" Width="400px" TextMode="MultiLine" Rows="5"></asp:textbox></FONT></TD>
												</TR>
												<TR vAlign="top">
													<TD class="header-item" vAlign="top" align="right" width="120" height="20"></TD>
													<TD class="item-text" vAlign="top" height="20">&nbsp;</TD>
												</TR>
												<TR>
													<TD class="header-item" vAlign="top" align="right" width="120" height="20"><br>
														<asp:label id="lblTel" runat="server">�B�z���ġG</asp:label></TD>
													<TD class="item-text" vAlign="top" height="20"><asp:textbox id="txtEffect" runat="server" Width="400px" TextMode="MultiLine" Rows="5"></asp:textbox></TD>
												</TR>
												<TR vAlign="top">
													<TD class="header-item" vAlign="top" align="right" width="120" height="20"></TD>
													<TD class="item-text" vAlign="top" height="20"><FONT face="�s�ө���"></FONT></TD>
												</TR>
												<TR>
													<TD class="header-item" style="HEIGHT: 35px" vAlign="top" align="right" width="120"
														height="35"><asp:label id="Label4" runat="server">������ơG</asp:label></TD>
													<TD class="item-text" style="HEIGHT: 35px" vAlign="top" height="35">
														<FIELDSET class="fieldset" style="WIDTH: 400px"><LEGEND align="right"><IMG alt="" src="images/icon01.gif" align="absBottom"><asp:hyperlink id="hlAttachment" runat="server">�s����[�ɮ�</asp:hyperlink></LEGEND><asp:label id="lblAttachment" runat="server">�S�����[�ɮ�</asp:label><asp:textbox id="txtFilename" runat="server" CssClass="Hide"></asp:textbox></FIELDSET>
													</TD>
												</TR>
												<TR>
													<TD class="header-item" vAlign="top" align="right" width="120" height="20">&nbsp;</TD>
													<TD class="item-text" vAlign="top" height="20">&nbsp;&nbsp;</TD>
												</TR>
												<TR vAlign="bottom">
													<TD class="header-item" vAlign="top" align="right" width="120" height="20">&nbsp;</TD>
													<TD class="item-text" vAlign="top" height="20"><FONT face="�s�ө���">&nbsp;</FONT></TD>
												</TR>
												<TR>
													<TD class="header-item" vAlign="top" align="right" width="120" height="20"></TD>
													<TD class="item-text" vAlign="top" height="20"></TD>
												</TR>
												<TR>
													<TD class="header-item" vAlign="top" align="right" width="120" height="20"><FONT face="�s�ө���"></FONT></TD>
													<TD class="item-text" vAlign="top" height="20"><FONT face="�s�ө���">&nbsp;</FONT></TD>
												</TR>
												<TR>
													<TD class="header-item" vAlign="top" align="right" width="120" height="20"><FONT face="�s�ө���"></FONT></TD>
													<TD class="item-text" vAlign="top" height="20"><IMG src="images/spacer.gif" width="8"></TD>
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
			<iframe name="nosee" width="0" height="0"></iframe></FONT>
	</body>
</HTML>
