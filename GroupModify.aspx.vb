Imports System
Imports System.Data
Imports System.Web.UI.WebControls
Imports Hsinyi.EIP.ProjectManagement.BusinessLogicLayer
Imports System.Text

Namespace Hsinyi.EIP.ProjectManagement.Web

    Public Class GroupModify
        Inherits System.Web.UI.Page

#Region " Web Form 設計工具產生的程式碼 "

        '此為 Web Form 設計工具所需的呼叫。
        <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

        End Sub
        Protected WithEvents Label1 As System.Web.UI.WebControls.Label
        Protected WithEvents Label4 As System.Web.UI.WebControls.Label
        Protected WithEvents userNamelbl As System.Web.UI.WebControls.Label
        Protected WithEvents ProjectList As System.Web.UI.HtmlControls.HtmlForm
        Protected WithEvents lbtnCancel As System.Web.UI.WebControls.LinkButton
        Protected WithEvents lbtnSave As System.Web.UI.WebControls.LinkButton
        Protected WithEvents ErrorMessage As System.Web.UI.WebControls.Label
        Protected WithEvents PanelModify As System.Web.UI.WebControls.Panel
        Protected WithEvents txtDescription As System.Web.UI.WebControls.TextBox
        Protected WithEvents txtName As System.Web.UI.WebControls.TextBox
        Protected WithEvents Members As System.Web.UI.WebControls.ListBox
        Protected WithEvents ProjectNameRequiredfieldvalidator As System.Web.UI.WebControls.RequiredFieldValidator

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
        Private _groupID As Integer
        Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
            If Request.QueryString("projectID") Is Nothing Then
                _projectID = 0
            Else
                _projectID = Convert.ToInt32(Request.QueryString("projectID"))
            End If

            If Request.QueryString("GroupID") Is Nothing Then
                _groupID = 0
            Else
                _groupID = Convert.ToInt32(Request.QueryString("GroupID"))
            End If

            '在這裡放置使用者程式碼以初始化網頁
            If Not Roles.isEnabledModifyProject(_projectID) Then
                Session("m_ErrorMessage") = "系統拒絕存取，您沒有權限存修改專案資訊。" '"您非本專案成員，因此無權限存取此專案資訊。"
                Response.Redirect("AlertMessage.aspx?Index=-1", True)
                PanelModify.Visible = False
            Else
                PanelModify.Visible = True
            End If



            If Not IsPostBack Then

                ' Load project with _projID when project id exists in the QueryString
                If _projectID <> 0 Then
                    BindMembers()
                Else
                    Session("m_ErrorMessage") = "系統錯誤，無法存取群組資訊。"  '"您非本專案成員，因此無權限存取此專案資訊。"
                    Response.Redirect("AlertMessage.aspx?Index=-1", True)
                End If

                If _groupID <> 0 Then
                    BindInfo()
                End If

            End If
        End Sub 'Page_Load

        Private Sub BindMembers()


            Members.DataSource = PMUser.GetMembers(_projectID)
            Members.DataValueField = "MemberID"
            Members.DataTextField = "Name"
            Members.DataBind()

        End Sub 'BindMembers
        '*********************************************************************
        '
        ' The BindInfo method retrieves the list of roles 
        ' and Load User information if user exists
        '
        '*********************************************************************

        Private Sub BindInfo()

            Dim group As New PMGroup(_groupID)
            group.Load()

            _projectID = group.ProjectID
            txtName.Text = group.Name
            txtDescription.Text = group.Description

            ' Sets the selected members of the project.
            Dim user As PMUser
            For Each user In group.Members
                Members.Items.FindByValue(user.MemberID.ToString()).Selected = True
            Next user

        End Sub 'BindInfo

        Private Sub lbtnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lbtnCancel.Click
            BackToUserPage()
        End Sub
        Private Sub BackToUserPage()
            Response.Redirect([String].Format("GroupDetail.aspx?ProjectID={0}&GroupID={1}&index=-1&projectIndex={2}", _projectID, _groupID, TabItem.ProjectTabIndex.ProjectGroup), True)
        End Sub

        Private Sub lbtnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lbtnSave.Click
            Dim group As New PMGroup(_groupID)
            group.Load()
            group.Name = txtName.Text
            group.Description = txtDescription.Text
            group.ProjectID = _projectID
            group.ModifiedUserID = PMSecurity.GetUserID

            ' Get list of members the user selected.
            Dim selectedMembers As New UsersCollection
            Dim li As ListItem
            For Each li In Members.Items
                If li.Selected Then
                    Dim user As New PMUser

                    user.MemberID = Convert.ToInt32(li.Value)
                    selectedMembers.Add(user)
                End If
            Next li


            group.Members = selectedMembers
            
            If group.Save() Then

                _groupID = group.GroupID
                BackToUserPage()
            Else
                Session("m_ErrorMessage") = "系統錯誤，無法儲存群組資訊。"  '"您非本專案成員，因此無權限存取此專案資訊。"
                Response.Redirect("AlertMessage.aspx?Index=-1", True)
            End If

        End Sub 'SaveButton_Click
    End Class
End Namespace


