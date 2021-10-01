<% option explicit%>
<% 'Singleton%>
<!-- #include virtual=/Class/clsStr.inc -->

<% ' Interface of clsThis %>
<!-- #include virtual=/Class/itfThis.inc -->

<% '  %>
<!-- #include virtual=/Class/clsTemplateHeader.inc -->
<!-- #include virtual=/Class/List/clsGroupsList.inc -->
<!-- #include virtual=/Class/Layout/clsTreeLayout.inc -->

	<iframe name="nosee" width=200 height=300></iframe>

	<script language=javascript>
		function toChangeColor(obj) {
			if (obj.className=="cr") {
				obj.className="cb"
				obj.style.cursor="default";
			}else{
				obj.className="cr"
				obj.style.cursor="hand";
			}
		}
		
		function toSelect(e) {
			 
				var obj=window.opener
				obj.toAddUser('<%=Request("ClientID")%>',e.id,e.innerText,1,e.sType)
				self.close();
			 
			
		}
		
		function toShowSub(obj,intID) {
			var strHtml;
			var strTemp=obj.src;
			var objSpan=document.all["s" + intID];
			if (strTemp.indexOf("plus")>=0) {
				obj.src="/icon/minus2.gif";
				if (objSpan.isGet) {
					objSpan.style.display="";
					return false
				}
				strHtml="<table cellspacing=0 cellpadding=0>";
				strHtml+="<tr><td><img src='/icon/empty.gif'><img src='/icon/Line3.gif'><img src='/icon/empty.gif'></td>";
				strHtml+="<td><font color=red>　請稍待......</font></td></tr>";
				strHtml+="</table>";
				objSpan.innerHTML= strHtml;
				objSpan.isGet=true;
				nosee.location="/App/range/getUsersTree.asp?ParentType=<%=this.ParentType%>&Parent_ID=" + intID;
			}else{
					obj.src="/icon/plus2.gif";
					objSpan.style.display="none";
			}
		}
	</script>
	
		
</body>
</html>

<% 

	class clsThis
		Private m_intOwnerType
		Private m_strTarget
		Private m_isUserOnly
		Private m_isGroupOnly
		Public Sub Class_Initialize()
			m_intOwnerType	= request("OwnerType")
			 
			if Request("isUserOnly") = 1 Then
				m_isUserOnly = KM_Tree_ClickDisabled
			Else
				m_isUserOnly = KM_Tree_Click
			End IF
			
			If Request("isGroupOnly") = 1 then
				m_isGroupOnly = KM_Tree_Item
			else
				m_isGroupOnly = KM_Tree_Top
			end if
			
			if m_intOwnerType="" then
				m_intOwnerType=1
			end if
			Main()
		End Sub
		
		Public Function Main()
			dim objHeader
			set objHeader = new TemplateHeader
			objHeader.AddScript("/script/JS_Tree.js")
			objHeader.showHeader
			set objHeader=nothing
			call ShowTree(m_intOwnerType)
		End Function
		
		Public Function ParentType()
			ParentType=m_intOwnerType
		end function

		
		Private Function ShowTree(intConstTreeItemType)
			dim sq,rs
			dim strTree
			strTree=getTreeItem(intConstTreeItemType)
%>			
			<table width=90% ><tr><td>
				<Table border=0 cellspacing=0 cellpadding=0 >
					<tr><td valign=bottom><img src="/icon/Logo.gif"></td><td>信誼基金會</td></tr>
				</table>
				<%=strTree%>
			</td></tr></table>
		
<%		End Function
	
		Private Function getTreeItem(intConstTreeItemType)
				dim objList
				dim objTree
				dim ar
				dim strTree
		
			select Case cint(intConstTreeItemType)
				Case KM_Group
					set objList=new GroupsList
				Case KM_Team
					set objList=new TeamsList
				Case KM_Position
					set objList=new PositionsList
				Case KM_GroupPosition
					set objList=new GroupPositionsList
				Case else
			End Select				

			ar=objList.getList(KM_Fundation,KM_Fundation)
			set objTree=new TreeLayout
			objTree.TreeList=ar
			objTree.ItemType=intConstTreeItemType
			objTree.TreeType=m_isGroupOnly	'KM_Tree_Top 人員  KM_Tree_Item部門
			objTree.isClick	= m_isUserOnly  '部門可選		KM_Tree_ClickDisabled 部門不可選
			strTree=objTree.getTree()
			set objList=nothing
			set objTree=nothing
			getTreeItem=strTree
			
		End Function 
		
		Public Sub Class_Terminate()
		End Sub
		
	end class
	
%>