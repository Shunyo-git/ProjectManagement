Imports System
Imports System.Data
Imports System.Web.UI.WebControls
Imports Hsinyi.EIP.ProjectManagement.BusinessLogicLayer

Namespace Hsinyi.EIP.ProjectManagement.Web


    Public Class UserDetail
        Inherits System.Web.UI.Page

#Region " Web Form 設計工具產生的程式碼 "

        '此為 Web Form 設計工具所需的呼叫。
        <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

        End Sub
        Protected WithEvents ProjectList As System.Web.UI.HtmlControls.HtmlForm
        Protected WithEvents lbtnEdit As System.Web.UI.WebControls.LinkButton
        Protected WithEvents lbtnDelete As System.Web.UI.WebControls.LinkButton
        Protected WithEvents PanelModify As System.Web.UI.WebControls.Panel
        Protected WithEvents userNamelbl As System.Web.UI.WebControls.Label
        Protected WithEvents lblName As System.Web.UI.WebControls.Label
        Protected WithEvents rolelbl As System.Web.UI.WebControls.Label
        Protected WithEvents lblRole As System.Web.UI.WebControls.Label
        Protected WithEvents Label1 As System.Web.UI.WebControls.Label
        Protected WithEvents lblGroup As System.Web.UI.WebControls.Label
        Protected WithEvents Label5 As System.Web.UI.WebControls.Label
        Protected WithEvents lblSourceUserDepartment As System.Web.UI.WebControls.Label
        Protected WithEvents passwordlbl As System.Web.UI.WebControls.Label
        Protected WithEvents lblSourceUserTitle As System.Web.UI.WebControls.Label
        Protected WithEvents lblTel As System.Web.UI.WebControls.Label
        Protected WithEvents lblConnects As System.Web.UI.WebControls.Label
        Protected WithEvents Label4 As System.Web.UI.WebControls.Label
        Protected WithEvents lblDescription As System.Web.UI.WebControls.Label

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
            ' Obtain project id from the QueryString.
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


            If Not Roles.isEnabledViewProject(_projectID) Then
                Session("m_ErrorMessage") = "系統拒絕存取，您沒有權限存取此專案資訊。" '"您非本專案成員，因此無權限存取此專案資訊。"
                Response.Redirect("AlertMessage.aspx?Index=-1", True)
            End If


            If Roles.isEnabledModifyProject(_projectID) Then
                PanelModify.Visible = True

            Else
                PanelModify.Visible = False
            End If





            If Not IsPostBack Then
                lbtnDelete.Attributes.Add("OnClick", "return chkConfirmDelete();")
                ' Load project with _projID when project id exists in the QueryString
                If _MemberID <> 0 And _projectID <> 0 Then
                    BindInfo()
                Else
                    Session("m_ErrorMessage") = "系統錯誤，無法存取成員資訊。"  '"您非本專案成員，因此無權限存取此專案資訊。"
                    Response.Redirect("AlertMessage.aspx?Index=-1", True)
                End If
            End If
        End Sub

        '*********************************************************************
        '
        ' The BindInfo method retrieves the list of roles 
        ' and Load User information if user exists
        '
        '*********************************************************************

        Private Sub BindInfo()
            Dim User As New PMUser(_MemberID)
            User.Load()
            _projectID = User.ProjectID
            lblName.Text = User.DisplayName

            lblRole.Text = User.RoleName
            lblGroup.Text = User.GroupName
            lblSourceUserDepartment.Text = User.SourceUserDepartment
            lblSourceUserTitle.Text = User.SourceUserTitle
            lblConnects.Text = User.Connects
            lblDescription.Text = User.Description.Replace(vbCrLf, "<BR>")

        End Sub 'BindInfo



        Private Sub lbtnEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lbtnEdit.Click
            Response.Redirect([String].Format("UserModify.aspx?projectID={0}&index=-1&projectIndex={1}&memberID={2}", _projectID, TabItem.ProjectTabIndex.ProjectUser, _MemberID), False)
        End Sub 'lbtnEdit_Click

        Private Sub lbtnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lbtnDelete.Click
            PMUser.RemoveMember(_MemberID)
            BackToProjectUser()
        End Sub

        Private Sub BackToProjectUser()
            Response.Redirect([String].Format("ProjectUser.aspx?projectID={0}&index=-1&projectIndex={1}", _projectID, TabItem.ProjectTabIndex.ProjectUser), True)
        End Sub


    End Class
End Namespace