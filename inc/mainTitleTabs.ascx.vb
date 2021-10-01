Imports System
Imports System.Collections
Imports EIPSysSecurity
Imports Hsinyi.EIP.ProjectManagement.BusinessLogicLayer


Namespace Hsinyi.EIP.ProjectManagement.Web

    Public Class mainTitleTabs
        Inherits System.Web.UI.UserControl


#Region " Web Form 設計工具產生的程式碼 "

        '此為 Web Form 設計工具所需的呼叫。
        <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

        End Sub
        Protected WithEvents tabs As System.Web.UI.WebControls.DataList
        Protected WithEvents lblCurrentUser As System.Web.UI.WebControls.Label

        '注意: 下列預留位置宣告是 Web Form 設計工具需要的項目。
        '請勿刪除或移動它。
        Private designerPlaceholderDeclaration As System.Object

        Private Sub Page_Init(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Init
            'CODEGEN: 此為 Web Form 設計工具所需的方法呼叫
            '請勿使用程式碼編輯器進行修改。
            InitializeComponent()
        End Sub

#End Region

        Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load


            lblCurrentUser.Text = PMSecurity.GetDisplayName & " ( " & Roles.GetRolesName(Convert.ToInt32(PMSecurity.GetUserRole)) & " )  "

            Dim tabItems As New ArrayList

            ' If Request.IsAuthenticated Then
            ' add the Log tab (all users can see this tab)
            tabItems.Add(New TabItem("我的首頁", "index.aspx?index=" & tabItems.Count))
            ' End If


            'If TTSecurity.IsInRole(TTUser.UserRoleAdminPMgr) Then
            tabItems.Add(New TabItem("專案列表", "ProjectList.aspx?index=" & tabItems.Count))
            ' tabItems.Add(New TabItem("最新文件", "NewDocument.aspx?index=" & tabItems.Count))
            ' tabItems.Add(New TabItem("最新話題", "NewTopic.aspx?index=" & tabItems.Count))
            ' End If

            '僅流程管理者可新增專案
            If PMSecurity.IsInRole(Roles.FlowManagr) Then
                tabItems.Add(New TabItem("新增專案", "ProjectModify.aspx?index=" & tabItems.Count))
                ' tabItems.Add(New TabItem("權限管理", "ProjectSecurity.aspx?index=" & tabItems.Count))
            End If
            tabs.DataSource = tabItems
            tabs.SelectedIndex = IIf(Request("index") Is Nothing, 0, Convert.ToInt32(Request("index")))
            tabs.DataBind()

        End Sub 'Page_Load
    End Class 'myProjectTabs
End Namespace 'Hsinyi.EIP.ProjectManagement.Web
