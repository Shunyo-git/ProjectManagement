Imports System
Imports System.Data
Imports System.Web.UI.WebControls
Imports Hsinyi.EIP.ProjectManagement.BusinessLogicLayer

Namespace Hsinyi.EIP.ProjectManagement.Web


Public Class UserModify
    Inherits System.Web.UI.Page

#Region " Web Form 設計工具產生的程式碼 "

    '此為 Web Form 設計工具所需的呼叫。
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

        End Sub
    Protected WithEvents ProjectList As System.Web.UI.HtmlControls.HtmlForm
    Protected WithEvents passwordlbl As System.Web.UI.WebControls.Label
        Protected WithEvents rolelbl As System.Web.UI.WebControls.Label
        Protected WithEvents userNamelbl As System.Web.UI.WebControls.Label
    Protected WithEvents txtDescription As System.Web.UI.WebControls.TextBox
    Protected WithEvents lblDescription As System.Web.UI.WebControls.Label
        Protected WithEvents lblSourceUserDepartment As System.Web.UI.WebControls.Label
    Protected WithEvents Label1 As System.Web.UI.WebControls.Label
        Protected WithEvents ErrorMessage As System.Web.UI.WebControls.Label
    Protected WithEvents PanelModify As System.Web.UI.WebControls.Panel
    Protected WithEvents lblTel As System.Web.UI.WebControls.Label
        Protected WithEvents lbtnCancel As System.Web.UI.WebControls.LinkButton
    Protected WithEvents lbtnSave As System.Web.UI.WebControls.LinkButton
        Protected WithEvents txtName As HsinYiMemberPicker.HsinYi.EIP.Security.MemberPicker.MemberPicker
        Protected WithEvents txtConnects As System.Web.UI.WebControls.TextBox
        Protected WithEvents txtSourceUserTitle As System.Web.UI.WebControls.TextBox
        Protected WithEvents ddlGroup As System.Web.UI.WebControls.DropDownList
        Protected WithEvents ddlRole As System.Web.UI.WebControls.DropDownList
        Protected WithEvents txtSourceUserDepartment2 As System.Web.UI.WebControls.TextBox
        Protected WithEvents txtSourceUserDepartment As HsinYiMemberPicker.HsinYi.EIP.Security.MemberPicker.MemberPicker

    '注意: 下列預留位置宣告是 Web Form 設計工具需要的項目。
    '請勿刪除或移動它。
    Private designerPlaceholderDeclaration As System.Object

    Private Sub Page_Init(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Init
        'CODEGEN: 此為 Web Form 設計工具所需的方法呼叫
        '請勿使用程式碼編輯器進行修改。
        InitializeComponent()
    End Sub

#End Region
        Private _projectID As Integer
        Private _MemberID As Integer
        Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
            If Request.QueryString("projectID") Is Nothing Then
                _projectID = 0
            Else
                _projectID = Convert.ToInt32(Request.QueryString("projectID"))
            End If

            If Request.QueryString("MemberID") Is Nothing Then
                _MemberID = 0
            Else
                _MemberID = Convert.ToInt32(Request.QueryString("MemberID"))
            End If

            '在這裡放置使用者程式碼以初始化網頁
            If Not Roles.isEnabledViewProject(_projectID) Then
                Session("m_ErrorMessage") = "系統拒絕存取，您沒有權限存修改專案資訊。" '"您非本專案成員，因此無權限存取此專案資訊。"
                Response.Redirect("AlertMessage.aspx?Index=-1", True)
                PanelModify.Visible = False
            Else
                PanelModify.Visible = True
            End If


            If Not IsPostBack Then

                ' Load project with _projID when project id exists in the QueryString
                If _projectID <> 0 Then
                    BindDefaultRoleGroup()
                Else
                    Session("m_ErrorMessage") = "系統錯誤，無法存取成員資訊。"  '"您非本專案成員，因此無權限存取此專案資訊。"
                    Response.Redirect("AlertMessage.aspx?Index=-1", True)
                End If

                If _MemberID <> 0 Then
                    BindInfo()
                End If
            End If
        End Sub 'Page_Load
        Private Sub BindDefaultRoleGroup()
            ddlRole.DataSource = Roles.GetProjectRoles
            ddlRole.DataTextField = "Name"
            ddlRole.DataValueField = "RoleID"
            ddlRole.DataBind()

            ddlGroup.DataSource = PMGroup.GetProjectGroups(_projectID)
            ddlGroup.DataTextField = "Name"
            ddlGroup.DataValueField = "GroupID"
            ddlGroup.DataBind()
        End Sub
        Private Sub BindInfo()
            Dim User As New PMUser(_MemberID)
            User.Load()

            txtName.Text = User.DisplayName
            txtName.WhoID = User.UserID

            txtName.WhoName = EIPSysSecurity.clsUser.getUserName(User.UserID)
            txtName.WhoType = "0"
            txtSourceUserDepartment.Text = User.SourceUserDepartment
            txtSourceUserTitle.Text = User.SourceUserTitle
            txtConnects.Text = User.Connects
            txtDescription.Text = User.Description

            If Not User.Role Is Nothing Then
                ddlRole.Items.FindByValue(User.Role).Selected = True

            End If
            If User.Group > 0 Then
                ddlGroup.Items.FindByValue(User.Group).Selected = True

            End If

        End Sub 'BindInfo

        Private Sub lbtnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lbtnCancel.Click
            BackToUserPage()

        End Sub

        Private Sub lbtnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lbtnSave.Click


            Dim isUserFound As Boolean = False
            Dim isUserActiveManager As Boolean = False

            ' Save the user and ask the user object to make sure the username is
            ' found in the user account source (NT SAM or Active Directory)
            'Usr.Load()

            If Strings.Trim(txtName.WhoID) = "" Then
                Session("m_ErrorMessage") = "姓名不得空白，無法儲存成員資訊。"  '"您非本專案成員，因此無權限存取此專案資訊。"
                Response.Redirect("AlertMessage.aspx?Index=-1", True)
            Else
                Dim Usr As New PMUser(_MemberID)
                Usr.ProjectID = _projectID
                Usr.UserID = Convert.ToInt32(txtName.WhoID)
                Usr.Role = ddlRole.SelectedItem.Value
                Usr.Group = ddlGroup.SelectedItem.Value
                Usr.SourceUserDepartment = txtSourceUserDepartment.Text
                Usr.SourceUserTitle = txtSourceUserTitle.Text
                Usr.Connects = txtConnects.Text
                Usr.Description = txtDescription.Text
                Usr.ModifiedUserID = PMSecurity.GetUserID
                If Usr.Save() Then
                    _MemberID = Usr.MemberID
                    BackToUserPage()
                Else
                    Session("m_ErrorMessage") = "系統錯誤，無法儲存成員資訊。"  '"您非本專案成員，因此無權限存取此專案資訊。"
                    Response.Redirect("AlertMessage.aspx?Index=-1", True)
                End If
            End If


            'If Not (Password.Text Is String.Empty) Then
            '    Usr.Password = TTSecurity.Encrypt(Password.Text)
            'End If

        End Sub
        Private Sub BackToUserPage()
            Response.Redirect([String].Format("UserDetail.aspx?ProjectID={0}&MemberID={1}&index=-1&projectIndex={2}", _projectID, _MemberID, TabItem.ProjectTabIndex.ProjectUser), True)
        End Sub
    End Class 'UserModify

End Namespace 'Hsinyi.EIP.ProjectManagement.Web
