<%@ Page Language="vb" AutoEventWireup="false" Codebehind="Attachment.aspx.vb" Inherits="HsinYi.EIP.ProjectManagement.Web.Attachment" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>管理附加檔案</title>
		<meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
		<meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
		<meta content="JavaScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
		<LINK href="css/Styles.css" type="text/css" rel="stylesheet">
		<style type="text/css">LEGEND { FONT: bold 12px tahoma, verdana, geneva, lucida, 'lucida grande', arial, helvetica, sans-serif; COLOR: #666666 }
		</style>
		<script type="text/javascript">
		<!--
		var SESSIONURL = "";
		var IMGDIR_MISC = "classic-images/misc";
		
		 
		function toClear(e) {
			//刪除檔案後,清除檔案顯示相關資料
			var obj=Form1.txtFilename;
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
		
		function SaveAttachment(){
		opener.document.getElementById("lblAttachment").innerHTML = document.getElementById("lblFileNameList").innerHTML ;
		opener.document.getElementById("txtFilename").value = document.getElementById("txtFilename").value ;
		
		}
	
	 -->
		</script>
	</HEAD>
	<body onload="self.focus()">
		<script type="text/javascript">
	<!--
	function verify_upload(formobj)
	{
		var haveupload = false;
		for (var i=0; i < formobj.elements.length; i++)
		{
			var elm = formobj.elements[i];
			if (elm.type == 'file')
			{
				if (elm.value != "")
				{
					haveupload = true;
				}
			}
		}

		if (haveupload)
		{
			//toggle_display('uploading');
			uploading.style.display='none';
			return true;
		}
		else
		{
			alert("請點選「瀏覽...」按鈕並選擇一個要上傳的文件。");
			return false;
		}
	}

	//var win=window.dialogArguments ;
	//var attachlist = win.document.getElementById("attachlist");
 
	 var attachlist = window.opener.document.all["attachlist"];
	 
	//attachlist.innerHTML = "<div style=\"margin:2px\"><img class=\"inlineimg\" src=\"classic-images/attach/jpg.gif\" alt=\"Yukio_05.jpg\" border=\"0\" /> <a href=\"attachment.php?attachmentid=12328&amp;stc=1\" target=\"_blank\">Yukio_05.jpg</a> (49.4 KB)</div>";
	 -->
		</script>
		<form id="Form1" method="post" encType="multipart/form-data" runat="server">
			<FONT face="新細明體">&nbsp;
				<asp:textbox id="txtFilename" runat="server" CssClass="Hide"></asp:textbox><asp:label id="lblHidden" runat="server"></asp:label><asp:textbox id="txtFileSize" runat="server" CssClass="hide"></asp:textbox>
				<table class="tab-table" id="Main" cellSpacing="1" cellPadding="3" width="500" align="center"
					border="0">
					<tr>
						<td style="FONT-SIZE: 11pt" align="left" bgColor="#d2d291"><span class="smallfont" style="FONT-SIZE: 10pt; FLOAT: right"><A onclick="self.close()" href="#">關閉視窗</A></span>
							管理附加檔案
						</td>
					</tr>
					<tr>
						<td class="table-top-border" align="center"><asp:panel id="PanelUpload" runat="server">
								<FIELDSET class="fieldset"><LEGEND>上傳文件</LEGEND>
									<TABLE cellSpacing="3" cellPadding="0" width="500" border="0">
										<TR>
											<TD colSpan="2"><FONT face="新細明體"></FONT></TD>
										</TR>
										<TR vAlign="bottom">
											<TD>選取要上傳的檔案:
												<BR>
												<INPUT id="File1" style="WIDTH: 392px; HEIGHT: 22px" type="file" size="46" name="File1"
													runat="server">
											</TD>
											<TD align="right"></TD>
										</TR>
										<TR>
											<TD colSpan="2"><FONT face="新細明體">
													<asp:button id="btnUpload" runat="server" CssClass="button" Text="上載" Width="70px"></asp:button>&nbsp;
													<asp:Button id="btnSave" runat="server" CssClass="button" Text="確定"></asp:Button></FONT></TD>
										</TR>
									</TABLE>
								</FIELDSET>
							</asp:panel><br>
							<asp:Panel id="PanelView" runat="server">
								<fieldset class="fieldset"><LEGEND>檢視目前的附加檔案</LEGEND>
									<DIV style="PADDING-RIGHT: 3px; PADDING-LEFT: 3px; PADDING-BOTTOM: 3px; PADDING-TOP: 3px">
										<asp:label id="lblFileNameList" runat="server"></asp:label></DIV>
								</fieldset>
							</asp:Panel>
							<br>
							&nbsp;
							<div id="uploading" style="MARGIN-TOP: 3px; DISPLAY: none" align="center"><strong>正在上傳文件 
									- 請稍候</strong>
							</div>
							<br>
							<div style="MARGIN-TOP: 3px">&nbsp;&nbsp;&nbsp;&nbsp;</div>
						</td>
					</tr>
				</table>
			</FONT>
		</form>
	</body>
</HTML>
