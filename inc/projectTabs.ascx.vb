Imports System
Imports System.Collections
Imports Hsinyi.EIP.ProjectManagement.BusinessLogicLayer
Imports EIPSysSecurity

Namespace Hsinyi.EIP.ProjectManagement.Web

    Public Class projectTabs
        Inherits System.Web.UI.UserControl

#Region " Web Form 設計工具產生的程式碼 "

        '此為 Web Form 設計工具所需的呼叫。
        <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

        End Sub
        Protected WithEvents tabs As System.Web.UI.WebControls.DataList
        Protected WithEvents dlProjectTabs As System.Web.UI.WebControls.DataList
        Protected WithEvents lblProjectName As System.Web.UI.WebControls.Label

        '注意: 下列預留位置宣告是 Web Form 設計工具需要的項目。
        '請勿刪除或移動它。
        Private designerPlaceholderDeclaration As System.Object

        Private Sub Page_Init(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Init
            'CODEGEN: 此為 Web Form 設計工具所需的方法呼叫
            '請勿使用程式碼編輯器進行修改。
            InitializeComponent()
        End Sub

#End Region
        Private projectID As Integer
        Private mainIndex As Integer


        Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

            mainIndex = IIf(Request("index") Is Nothing, 0, Convert.ToInt32(Request("index")))
            projectID = IIf(Request("projectID") Is Nothing, 0, Convert.ToInt32(Request("projectID")))

            If Request.QueryString("projectID") Is Nothing Then
                projectID = 0
            End If

            'If Not Page.IsPostBack Then
            If projectID <> 0 Then
                AddTabItems()
                Dim pj As New Project(projectID)
                pj.Load()
                lblProjectName.Text = "目前專案：" & pj.ProjectName
                'lblProjectName.Text += "<BR>" & pj.StartDate.ToShortDateString & "~" & pj.EndDate.ToShortDateString

            End If
            'End If




        End Sub 'Page_Load

        Private Sub AddTabItems()
            Dim tabItems As New ArrayList

            ' Always add the project list tab.  If the user was able to see the Administration tab, they'll
            ' always be able to view this tab.

            tabItems.Add(New TabItem("專案首頁", "ProjectDetails.aspx?index=" & mainIndex & "&projectID=" & projectID & "&projectIndex=" & tabItems.Count))
            tabItems.Add(New TabItem("專案進度表", "ProjectDocument.aspx?index=" & mainIndex & "&projectID=" & projectID & "&projectIndex=" & tabItems.Count & "&categoryIndex=" & Convert.ToInt32(TabItem.CategoryTabIndex.Timeline)))

            tabItems.Add(New TabItem("專案組別", "ProjectGroup.aspx?index=" & mainIndex & "&projectID=" & projectID & "&projectIndex=" & tabItems.Count))
            tabItems.Add(New TabItem("人力配置", "ProjectUser.aspx?index=" & mainIndex & "&projectID=" & projectID & "&projectIndex=" & tabItems.Count))
            tabItems.Add(New TabItem("提醒/通知", "ProjectCalendar.aspx?index=" & mainIndex & "&projectID=" & projectID & "&projectIndex=" & tabItems.Count))
            tabItems.Add(New TabItem("特殊事件記錄", "ProjectEvent.aspx?index=" & mainIndex & "&projectID=" & projectID & "&projectIndex=" & tabItems.Count))
            tabItems.Add(New TabItem("專案文件", "ProjectDocument.aspx?index=" & mainIndex & "&projectID=" & projectID & "&projectIndex=" & tabItems.Count))
            tabItems.Add(New TabItem("討論區", "ProjectForum.aspx?index=" & mainIndex & "&projectID=" & projectID & "&projectIndex=" & tabItems.Count))
            tabItems.Add(New TabItem("連絡表", "ProjectContact.aspx?index=" & mainIndex & "&projectID=" & projectID & "&projectIndex=" & tabItems.Count))

            '僅專案負責人即只要執行者可新增人力
            'If PMSecurity.IsInRole(PMUser.FlowManagr) Then
            '    tabItems.Add(New TabItem("新增專案", "ProjectNew.aspx?index=" & mainIndex & "&adminIndex=" & tabItems.Count))
            '    tabItems.Add(New TabItem("權限管理", "ProjectNew.aspx?index=" & mainIndex & "&adminIndex=" & tabItems.Count))
            'End If



            dlProjectTabs.DataSource = tabItems
            dlProjectTabs.SelectedIndex = IIf(Request("projectIndex") Is Nothing, 0, Convert.ToInt16(Request("projectIndex")))
            dlProjectTabs.DataBind()
        End Sub
    End Class
End Namespace

