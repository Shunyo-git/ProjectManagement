Imports System
Imports System.Web.UI.WebControls
Imports Hsinyi.EIP.ProjectManagement.BusinessLogicLayer
Imports Hsinyi.EIP.ProjectManagement.DataAccessLayer

Namespace Hsinyi.EIP.ProjectManagement.Web

    Public Class ProjectDetails
        Inherits System.Web.UI.Page

#Region " Web Form 設計工具產生的程式碼 "

        '此為 Web Form 設計工具所需的呼叫。
        <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

        End Sub
        Protected WithEvents ProjectList As System.Web.UI.HtmlControls.HtmlForm
        Protected WithEvents lblState As System.Web.UI.WebControls.Label
        Protected WithEvents lblType As System.Web.UI.WebControls.Label
        Protected WithEvents lblManager As System.Web.UI.WebControls.Label
        Protected WithEvents lblDescription As System.Web.UI.WebControls.Label
        Protected WithEvents lblSupplementaryUnits As System.Web.UI.WebControls.Label
        Protected WithEvents lblSupplementarySubsidy As System.Web.UI.WebControls.Label
        Protected WithEvents lblAssistUnits As System.Web.UI.WebControls.Label
        Protected WithEvents lblProjectName As System.Web.UI.WebControls.Label
        Protected WithEvents Label1 As System.Web.UI.WebControls.Label
        Protected WithEvents lblManagerDepartment As System.Web.UI.WebControls.Label
        Protected WithEvents lblChiefDepartment As System.Web.UI.WebControls.Label
        Protected WithEvents Label3 As System.Web.UI.WebControls.Label
        Protected WithEvents lblChief As System.Web.UI.WebControls.Label
        Protected WithEvents lblProjectTime As System.Web.UI.WebControls.Label
        Protected WithEvents PanelModify As System.Web.UI.WebControls.Panel
        Protected WithEvents lbtnEdit As System.Web.UI.WebControls.LinkButton
        Protected WithEvents lbtnDelete As System.Web.UI.WebControls.LinkButton
        Protected WithEvents PanelView As System.Web.UI.WebControls.Panel
        Protected WithEvents ErrorMessage As System.Web.UI.WebControls.Label

        '注意: 下列預留位置宣告是 Web Form 設計工具需要的項目。
        '請勿刪除或移動它。
        Private designerPlaceholderDeclaration As System.Object

        Private Sub Page_Init(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Init
            'CODEGEN: 此為 Web Form 設計工具所需的方法呼叫
            '請勿使用程式碼編輯器進行修改。
            InitializeComponent()
        End Sub

#End Region
        ' Contains the id of current project
        Private _projectID As Integer
        Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
            lbtnDelete.Attributes.Add("OnClick", "return chkConfirmDelete();")

            '在這裡放置使用者程式碼以初始化網頁
            If Request.QueryString("projectID") Is Nothing Then
                _projectID = 0
            Else
                _projectID = Convert.ToInt32(Request.QueryString("projectID"))
            End If

            'Page.Trace.IsEnabled = True
            'Response.Write(PMSecurity.IsInRole(Roles.UserEnabledViewProject))
            'Response.Write(PMUser.IsProjectMember(_projectID, PMSecurity.GetUserID))
            'Exit Sub

            If Not Roles.isEnabledViewProject(_projectID) Then
                Session("m_ErrorMessage") = "系統拒絕存取，您沒有權限存取此專案資訊。" '"您非本專案成員，因此無權限存取此專案資訊。"
                Response.Redirect("AlertMessage.aspx?Index=-1", True)
            End If

            If Roles.isEnabledModifyProject(_projectID) Then
                PanelModify.Visible = True
            Else
                PanelModify.Visible = False
            End If

            ' Obtain project id from the QueryString.

            If Not IsPostBack Then

                ' Load project with _projID when project id exists in the QueryString
                If _projectID <> 0 Then
                    BindProject()
                Else
                    Session("m_ErrorMessage") = "系統錯誤，無法存取專案資訊。"  '"您非本專案成員，因此無權限存取此專案資訊。"
                    Response.Redirect("AlertMessage.aspx?Index=-1", True)
                End If
            End If
        End Sub
        '*********************************************************************
        '
        ' The BindProject method retrieves the project with _projID and populates
        ' the web controls with the project's info.
        '
        '*********************************************************************

        Private Sub BindProject()

            Dim project As New project(_projectID)
            project.Load()

            Dim manager As New PMUser(project.ManagerUserID, _projectID)
            manager.LoadProjectMember()
            Dim Chief As New PMUser(project.ChiefUserID, _projectID)
            Chief.LoadProjectMember()

            ' Populate web controls with the project's info. ( for View )
            lblProjectName.Text = project.ProjectName
            lblDescription.Text = IIf(project.ProjectDescription Is String.Empty, "", project.ProjectDescription.Replace(vbCrLf, "<BR>"))
            lblManager.Text = manager.DisplayName
            lblManagerDepartment.Text = manager.SourceUserDepartment
            lblChief.Text = Chief.DisplayName
            lblChiefDepartment.Text = Chief.SourceUserDepartment
            lblProjectTime.Text = project.StartDate & " ~ " & project.EndDate
            lblState.Text = project.StateName
            lblState.ForeColor = ProjectState.ShowStateColor(project.StateID)
            lblType.Text = project.TypeName
            lblSupplementarySubsidy.Text = IIf(project.SupplementarySubsidy Is String.Empty, "", project.SupplementarySubsidy.Replace(vbCrLf, "<BR>"))
            lblSupplementaryUnits.Text = IIf(project.SupplementaryUnits Is String.Empty, "", project.SupplementaryUnits.Replace(vbCrLf, "<BR>"))
            lblAssistUnits.Text = IIf(project.AssistUnits Is String.Empty, "", project.AssistUnits.Replace(vbCrLf, "<BR>"))

            manager = Nothing
            Chief = Nothing

            ' Populate web controls with the project's info. ( for Modify)

            'Managers.Items.FindByValue(project.ManagerUserID.ToString()).Selected = True

            ' Sets the selected members of the project.
            'Dim user As PMUser
            'For Each user In project.Members
            '    'Members.Items.FindByValue(user.UserID.ToString()).Selected = True
            'Next user

            ' BindCategoriesGrid(project.Categories)
        End Sub 'BindProject



        Private Sub lbtnEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lbtnEdit.Click

            Response.Redirect([String].Format("ProjectModify.aspx?projectID={0}&index=-1&projectIndex={1}", _projectID, TabItem.ProjectTabIndex.ProjectDetail), False)
        End Sub 'lbtnEdit_Click



        Private Sub lbtnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lbtnDelete.Click
            Project.Remove(_projectID)
        End Sub
    End Class


End Namespace